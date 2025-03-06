VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Kasa Wire Labels"
   ClientHeight    =   2340
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5580
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1615.108
   ScaleMode       =   0  'User
   ScaleWidth      =   5239.909
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1950
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1950
      ScaleWidth      =   2550
      TabIndex        =   3
      ToolTipText     =   "Kasa Industrial Controls, Inc."
      Top             =   240
      Width           =   2550
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2880
      TabIndex        =   0
      ToolTipText     =   "Close the About Window"
      Top             =   1920
      Width           =   2580
   End
   Begin VB.Label lblFormatLocked 
      Alignment       =   2  'Center
      Caption         =   "Format Locked"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   2880
      TabIndex        =   7
      Top             =   720
      Width           =   2565
   End
   Begin VB.Label lblKasa 
      Alignment       =   2  'Center
      Caption         =   "418 East Avenue B."
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   6
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblKasa 
      Alignment       =   2  'Center
      Caption         =   "Salina, KS  67401"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label lblKasa 
      Alignment       =   2  'Center
      Caption         =   "Kasa Industrial Controls"
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2850
      TabIndex        =   1
      Top             =   120
      Width           =   2565
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Version"
      Height          =   225
      Left            =   2850
      TabIndex        =   2
      Top             =   360
      Width           =   2565
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    ' Unload the about form.
    Unload Me
End Sub

Private Sub Form_Load()

    '200510 N.F.
    'Set form caption
    Me.Caption = GetFormCaption(Me.Name)
    'Assume "About xxxxxx"
    If Left(Me.Caption, 6) = "About " Then
        lblTitle.Caption = Right(Me.Caption, Len(Me.Caption) - 6)
    Else
        'Don't know what we have
        Debug.Assert False
        lblTitle.Caption = ""
    End If
    
    lblFormatLocked.Visible = blnLockFormat
    
    ' Set the version label to the current version.
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub picLogo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And (Shift = vbShiftMask + vbCtrlMask) Then
        blnLockFormat = Not blnLockFormat
        lblFormatLocked.Visible = blnLockFormat
    End If
End Sub
