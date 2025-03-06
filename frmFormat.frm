VERSION 5.00
Begin VB.Form frmFormat 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Label Format"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   8055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPaste 
      Caption         =   "0."
      Enabled         =   0   'False
      Height          =   255
      Left            =   960
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer timCalc 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   6600
   End
   Begin VB.OptionButton optWire 
      Caption         =   "Wire Labels"
      Height          =   375
      Left            =   240
      TabIndex        =   36
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton optTerminal 
      Caption         =   "Terminal Labels"
      Height          =   375
      Left            =   1920
      TabIndex        =   35
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame fraOptions 
      Height          =   3735
      Left            =   4680
      TabIndex        =   23
      Top             =   2520
      Width           =   2895
      Begin VB.CheckBox chkAutoSize 
         Caption         =   "Automatically Size Font To Fit"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox chkOptical 
         Caption         =   "Optically Sensed Labels"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   2055
      End
      Begin VB.ListBox lstFormats 
         Height          =   2010
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lblFormat 
         Caption         =   "Format Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculator"
      DisabledPicture =   "frmFormat.frx":0000
      Height          =   615
      Left            =   1800
      Picture         =   "frmFormat.frx":06E2
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   5400
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   615
      Left            =   3600
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Frame fraWire 
      Caption         =   "Wire Label Formats"
      Height          =   5775
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   7575
      Begin VB.TextBox txtLines 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6240
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtLabelsPerRow 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6240
         TabIndex        =   11
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtSpacingTB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtTop 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtSpacingRL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Caption         =   "Currently Selected Option:"
         Height          =   255
         Left            =   4440
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblLines 
         Alignment       =   1  'Right Justify
         Caption         =   "Printed Lines Per Label:"
         Height          =   255
         Left            =   4440
         TabIndex        =   16
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblLabelsPerRow 
         Alignment       =   1  'Right Justify
         Caption         =   "Labels Per Row:"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblOption 
         AutoSize        =   -1  'True
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4560
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Line linSpacingTB 
         Index           =   1
         X1              =   3360
         X2              =   3840
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line linSpacingTB 
         Index           =   6
         X1              =   3360
         X2              =   3840
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line linSpacingTB 
         Index           =   0
         X1              =   3600
         X2              =   3600
         Y1              =   2760
         Y2              =   720
      End
      Begin VB.Line linSpacingTB 
         Index           =   2
         X1              =   3720
         X2              =   3600
         Y1              =   840
         Y2              =   720
      End
      Begin VB.Line linSpacingTB 
         Index           =   3
         X1              =   3480
         X2              =   3600
         Y1              =   840
         Y2              =   720
      End
      Begin VB.Line linSpacingTB 
         Index           =   5
         X1              =   3480
         X2              =   3600
         Y1              =   2640
         Y2              =   2760
      End
      Begin VB.Line linSpacingTB 
         Index           =   4
         X1              =   3720
         X2              =   3600
         Y1              =   2640
         Y2              =   2760
      End
      Begin VB.Shape shpLabelClear 
         Height          =   1335
         Index           =   0
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   735
      End
      Begin VB.Shape shpLabelWhite 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   0
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   735
      End
      Begin VB.Shape shpLabelClear 
         Height          =   1335
         Index           =   3
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   2760
         Width           =   735
      End
      Begin VB.Shape shpLabelClear 
         Height          =   1335
         Index           =   2
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   2760
         Width           =   735
      End
      Begin VB.Line linTopMar 
         Index           =   0
         X1              =   2640
         X2              =   2640
         Y1              =   120
         Y2              =   720
      End
      Begin VB.Line linTopMar 
         Index           =   4
         X1              =   2760
         X2              =   2640
         Y1              =   240
         Y2              =   120
      End
      Begin VB.Line linTopMar 
         Index           =   3
         X1              =   2520
         X2              =   2640
         Y1              =   240
         Y2              =   120
      End
      Begin VB.Line linTopMar 
         Index           =   2
         X1              =   2520
         X2              =   2640
         Y1              =   600
         Y2              =   720
      End
      Begin VB.Line linTopMar 
         Index           =   1
         X1              =   2760
         X2              =   2640
         Y1              =   600
         Y2              =   720
      End
      Begin VB.Line linSpacingRL 
         Index           =   0
         X1              =   960
         X2              =   2400
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line linSpacingRL 
         Index           =   3
         X1              =   1080
         X2              =   960
         Y1              =   2280
         Y2              =   2400
      End
      Begin VB.Line linSpacingRL 
         Index           =   4
         X1              =   1080
         X2              =   960
         Y1              =   2520
         Y2              =   2400
      End
      Begin VB.Line linSpacingRL 
         Index           =   2
         X1              =   2280
         X2              =   2400
         Y1              =   2280
         Y2              =   2400
      End
      Begin VB.Line linSpacingRL 
         Index           =   1
         X1              =   2280
         X2              =   2400
         Y1              =   2520
         Y2              =   2400
      End
      Begin VB.Line linLeftMar 
         Index           =   0
         X1              =   0
         X2              =   960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line linLeftMar 
         Index           =   3
         X1              =   120
         X2              =   0
         Y1              =   960
         Y2              =   1080
      End
      Begin VB.Line linLeftMar 
         Index           =   4
         X1              =   120
         X2              =   0
         Y1              =   1200
         Y2              =   1080
      End
      Begin VB.Line linLeftMar 
         Index           =   1
         X1              =   840
         X2              =   960
         Y1              =   960
         Y2              =   1080
      End
      Begin VB.Line linLeftMar 
         Index           =   2
         X1              =   840
         X2              =   960
         Y1              =   1200
         Y2              =   1080
      End
      Begin VB.Shape shpLabelWhite 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   2
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   2760
         Width           =   735
      End
      Begin VB.Shape shpLabelWhite 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   3
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   2760
         Width           =   735
      End
      Begin VB.Line linWidth 
         Index           =   6
         X1              =   960
         X2              =   960
         Y1              =   4200
         Y2              =   4680
      End
      Begin VB.Line linWidth 
         Index           =   0
         X1              =   1680
         X2              =   1680
         Y1              =   4200
         Y2              =   4680
      End
      Begin VB.Line linWidth 
         Index           =   3
         X1              =   960
         X2              =   1680
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line linWidth 
         Index           =   5
         X1              =   1080
         X2              =   960
         Y1              =   4560
         Y2              =   4440
      End
      Begin VB.Line linWidth 
         Index           =   4
         X1              =   1080
         X2              =   960
         Y1              =   4320
         Y2              =   4440
      End
      Begin VB.Line linWidth 
         Index           =   2
         X1              =   1560
         X2              =   1680
         Y1              =   4320
         Y2              =   4440
      End
      Begin VB.Line linWidth 
         Index           =   1
         X1              =   1560
         X2              =   1680
         Y1              =   4560
         Y2              =   4440
      End
      Begin VB.Line linLeftMar 
         Index           =   6
         X1              =   960
         X2              =   960
         Y1              =   840
         Y2              =   1320
      End
      Begin VB.Line linTopMar 
         Index           =   6
         X1              =   2400
         X2              =   2880
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line linSpacingRL 
         Index           =   6
         X1              =   960
         X2              =   960
         Y1              =   2160
         Y2              =   2640
      End
      Begin VB.Line linSpacingRL 
         Index           =   5
         X1              =   2400
         X2              =   2400
         Y1              =   2160
         Y2              =   2640
      End
      Begin VB.Line linHeight 
         Index           =   0
         X1              =   360
         X2              =   840
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line linHeight 
         Index           =   1
         X1              =   360
         X2              =   840
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line linHeight 
         Index           =   4
         X1              =   600
         X2              =   600
         Y1              =   2760
         Y2              =   3240
      End
      Begin VB.Line linHeight 
         Index           =   5
         X1              =   720
         X2              =   600
         Y1              =   2880
         Y2              =   2760
      End
      Begin VB.Line linHeight 
         Index           =   6
         X1              =   480
         X2              =   600
         Y1              =   2880
         Y2              =   2760
      End
      Begin VB.Line linHeight 
         Index           =   3
         X1              =   480
         X2              =   600
         Y1              =   3120
         Y2              =   3240
      End
      Begin VB.Line linHeight 
         Index           =   2
         X1              =   720
         X2              =   600
         Y1              =   3120
         Y2              =   3240
      End
      Begin VB.Line linTopMar 
         Index           =   5
         X1              =   2400
         X2              =   2880
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line linLeftMar 
         Index           =   5
         X1              =   0
         X2              =   0
         Y1              =   840
         Y2              =   1320
      End
      Begin VB.Label lblDimensions 
         Caption         =   "Enter All Dimensions In Inches"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   5160
         Width           =   2175
      End
      Begin VB.Shape shpLabelClear 
         Height          =   1335
         Index           =   1
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   735
      End
      Begin VB.Shape shpLabelWhite 
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   1
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame fraTerminal 
      Caption         =   "Terminal Label Formats"
      Height          =   5775
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   7575
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   360
         Picture         =   "frmFormat.frx":09EC
         ScaleHeight     =   5175
         ScaleWidth      =   3975
         TabIndex        =   18
         Top             =   480
         Width           =   3975
         Begin VB.TextBox txtTSpacing 
            Height          =   285
            Left            =   3000
            TabIndex        =   22
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtTHeight 
            Height          =   285
            Left            =   240
            TabIndex        =   21
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtTLeft 
            Height          =   285
            Left            =   240
            TabIndex        =   19
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox txtTWidth 
            Height          =   285
            Left            =   1620
            TabIndex        =   20
            Top             =   4440
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "LINE 3 PRINTS ON TWO LINES"
            Height          =   375
            Left            =   1320
            TabIndex        =   34
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "LINE 2"
            Height          =   195
            Left            =   1680
            TabIndex        =   32
            Top             =   1600
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "LINE 1"
            Height          =   195
            Left            =   1680
            TabIndex        =   31
            Top             =   600
            Width           =   615
         End
         Begin VB.Line lneHeight 
            Index           =   3
            X1              =   360
            X2              =   360
            Y1              =   960
            Y2              =   480
         End
         Begin VB.Line lneSpacing 
            Index           =   7
            X1              =   3600
            X2              =   3480
            Y1              =   840
            Y2              =   720
         End
         Begin VB.Line lneSpacing 
            Index           =   6
            X1              =   3480
            X2              =   3360
            Y1              =   1680
            Y2              =   1560
         End
         Begin VB.Line lneSpacing 
            Index           =   5
            X1              =   3600
            X2              =   3480
            Y1              =   1560
            Y2              =   1680
         End
         Begin VB.Line lneSpacing 
            Index           =   4
            X1              =   3480
            X2              =   3360
            Y1              =   720
            Y2              =   840
         End
         Begin VB.Line lneSpacing 
            Index           =   3
            X1              =   3480
            X2              =   3480
            Y1              =   1440
            Y2              =   1680
         End
         Begin VB.Line lneSpacing 
            Index           =   2
            X1              =   3480
            X2              =   3480
            Y1              =   720
            Y2              =   960
         End
         Begin VB.Line lneSpacing 
            Index           =   1
            X1              =   2880
            X2              =   3720
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line lneSpacing 
            Index           =   0
            X1              =   2880
            X2              =   3720
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line lneLeftMar 
            Index           =   4
            X1              =   145
            X2              =   25
            Y1              =   2760
            Y2              =   2640
         End
         Begin VB.Line lneLeftMar 
            Index           =   3
            X1              =   1225
            X2              =   1105
            Y1              =   2640
            Y2              =   2520
         End
         Begin VB.Line lneLeftMar 
            Index           =   2
            X1              =   145
            X2              =   25
            Y1              =   2520
            Y2              =   2640
         End
         Begin VB.Line lneLeftMar 
            Index           =   1
            X1              =   1225
            X2              =   1105
            Y1              =   2640
            Y2              =   2760
         End
         Begin VB.Line lneLeftMar 
            Index           =   0
            X1              =   1225
            X2              =   25
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line lneHeight 
            Index           =   7
            X1              =   480
            X2              =   360
            Y1              =   840
            Y2              =   960
         End
         Begin VB.Line lneHeight 
            Index           =   6
            X1              =   240
            X2              =   360
            Y1              =   600
            Y2              =   480
         End
         Begin VB.Line lneHeight 
            Index           =   5
            X1              =   360
            X2              =   240
            Y1              =   960
            Y2              =   840
         End
         Begin VB.Line lneHeight 
            Index           =   4
            X1              =   480
            X2              =   360
            Y1              =   600
            Y2              =   480
         End
         Begin VB.Line lneHeight 
            Index           =   1
            X1              =   1080
            X2              =   120
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line lneHeight 
            Index           =   0
            X1              =   1080
            X2              =   120
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line lneWidth 
            Index           =   8
            X1              =   2520
            X2              =   2715
            Y1              =   4680
            Y2              =   4560
         End
         Begin VB.Line lneWidth 
            Index           =   7
            X1              =   2520
            X2              =   2715
            Y1              =   4440
            Y2              =   4560
         End
         Begin VB.Line lneWidth 
            Index           =   5
            X1              =   1245
            X2              =   1440
            Y1              =   4560
            Y2              =   4440
         End
         Begin VB.Line lneWidth 
            Index           =   4
            X1              =   1245
            X2              =   1440
            Y1              =   4560
            Y2              =   4680
         End
         Begin VB.Line lneWidth 
            Index           =   3
            X1              =   2400
            X2              =   2710
            Y1              =   4560
            Y2              =   4560
         End
         Begin VB.Line lneWidth 
            Index           =   2
            X1              =   1250
            X2              =   1560
            Y1              =   4560
            Y2              =   4560
         End
         Begin VB.Line lneWidth 
            Index           =   1
            X1              =   2695
            X2              =   2695
            Y1              =   4080
            Y2              =   4800
         End
         Begin VB.Line lneWidth 
            Index           =   0
            X1              =   1250
            X2              =   1250
            Y1              =   4080
            Y2              =   4800
         End
         Begin VB.Shape shpLabelWhite 
            BackStyle       =   1  'Opaque
            Height          =   495
            Index           =   7
            Left            =   1250
            Shape           =   4  'Rounded Rectangle
            Top             =   3450
            Width           =   1470
         End
         Begin VB.Shape shpLabelWhite 
            BackStyle       =   1  'Opaque
            Height          =   495
            Index           =   6
            Left            =   1250
            Shape           =   4  'Rounded Rectangle
            Top             =   2450
            Width           =   1470
         End
         Begin VB.Shape shpLabelWhite 
            BackStyle       =   1  'Opaque
            Height          =   495
            Index           =   5
            Left            =   1250
            Shape           =   4  'Rounded Rectangle
            Top             =   1450
            Width           =   1475
         End
         Begin VB.Shape shpLabelWhite 
            BackStyle       =   1  'Opaque
            Height          =   495
            Index           =   4
            Left            =   1250
            Shape           =   4  'Rounded Rectangle
            Top             =   450
            Width           =   1475
         End
      End
      Begin VB.Label lblTOption 
         AutoSize        =   -1  'True
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4560
         TabIndex        =   30
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label1 
         Caption         =   "Currently Selected Option:"
         Height          =   375
         Left            =   4440
         TabIndex        =   29
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "LINE 2"
      Height          =   195
      Left            =   2280
      TabIndex        =   33
      Top             =   4200
      Width           =   495
   End
   Begin VB.Line lneHeight 
      Index           =   2
      X1              =   960
      X2              =   960
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line lneWidth 
      Index           =   6
      X1              =   3120
      X2              =   3315
      Y1              =   6240
      Y2              =   6120
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Format..."
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Format..."
      End
   End
End
Attribute VB_Name = "frmFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_PATH = 260
Private lngWindow As Long
Private lngCalc As Long
Private intLeft As Long
Private intTop As Long
Private lngProc As Long
Private lngExit As Long
Private lngLength As Long


' Offsets for the copy button
' Since the controls are in a frame, the paste button
' requires a coordinate offset so it aligns properly
Private Const H_OFFSET = 260    ' Horizontal Offset
Private Const V_OFFSET = 600    ' Vertical Offset
Private Const TH_OFFSET = 335   ' Additional Terminal Frame Horizontal Offset
Private Const TV_OFFSET = 515   ' Additional Terminal Frame Vertical Offset

Private lngVOffset As Long

' Declare all the windows API functions we will be using.
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE

Private txtSelected As TextBox

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Sub chkAutoSize_Click()
' This subroutine sets the format description to "New Format" when
' the Auto Size check box value changes.

    txtDescription.Text = "New Format"
End Sub

Private Sub chkAutoSize_GotFocus()
' This subroutine hides the paste command button and changes
' the currently selected option label text to "Optically Sensed Labels"

    cmdPaste.Visible = False
    
    If optWire.Value = True Then
        lblOption.Caption = "Automatically Size Font To Fit"
    End If
    
    If optTerminal.Value = True Then
        lblOption.Caption = "Automatically Size Font To Fit"
    End If
    
End Sub

Private Sub chkAutoSize_KeyPress(KeyAscii As Integer)
' This subroutine changes the return key to the tab key when
' the chkOptical check box has focus.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub chkOptical_Click()
' This subroutine sets the format description to "New Format" when
' the optically sensed check box value changes.

    txtDescription.Text = "New Format"
End Sub

Private Sub chkOptical_GotFocus()
' This subroutine hides the paste command button and changes
' the currently selected option label text to "Optically Sensed Labels"

    cmdPaste.Visible = False
    
    If optWire.Value = True Then
        lblOption.Caption = "Optically Sensed Labels"
    End If
    
    If optTerminal.Value = True Then
        lblTOption.Caption = "Optically Sensed Labels"
    End If
    
End Sub

Private Sub chkOptical_KeyPress(KeyAscii As Integer)
' This subroutine changes the return key to the tab key when
' the chkOptical check box has focus.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub cmdCalc_Click()
' This subroutine runs the calculator program and initializes
' the variables to read directly from the calculator text box.

    Dim dblShell As Double
    Dim lngShell As Long
    Dim rectCalc As RECT
    Dim intTemp1 As Long
    Dim intTemp2 As Long
    Dim strTemp As String
    
    ' Run the calculator program built in to Windows
    lngShell = Shell(GetWinPath & "\system32\CALC.EXE", vbNormalFocus)
    
    ' Find the process ID of the calculator program we just ran
    lngProc = OpenProcess(&H400, False, lngShell)
    
    ' Find the code to exit the process of the calculator
    Call GetExitCodeProcess(lngProc, lngExit)
    
    ' Find the window handle to the Calculator program
    lngCalc = FindWindow(vbNullString, "Calculator")
    
    ' Change the text of the calculator window to "Kasa Wire Labels Calculator"
    Call SetWindowText(lngCalc, "Kasa Wire Labels Calculator")
    
    ' Get the rectangular coordinates of the calculator program
    Call GetWindowRect(lngCalc, rectCalc)
    
    ' Set the calculator to always on top
    Call SetWindowPos(lngCalc, -1, rectCalc.Left, rectCalc.Top, (rectCalc.Right - rectCalc.Left), (rectCalc.Bottom - rectCalc.Top), 0)
    
    ' Find the handle to the calculator's text box
    
    '200510 N.F.
    'A value of y = 66 covers all scenarios except XP extra large
    lngWindow = WindowFromPoint(rectCalc.Left + 232, rectCalc.Top + 66)
    If lngWindow = lngCalc Then
        lngWindow = WindowFromPoint(rectCalc.Left + 232, rectCalc.Top + 75)
    End If
    
    ' Start the timer and make sure the paste button is visible and enabled
    timCalc.Enabled = True
    cmdPaste.Enabled = True
    cmdPaste.Visible = True
    cmdCalc.Enabled = False
End Sub

Private Sub cmdCancel_Click()
' This subroutine closes the label format dialog box.

    Unload frmFormat
End Sub

Private Sub cmdCancel_GotFocus()
' This subroutine hides the paste button and the current and option labels

    cmdPaste.Visible = False
    lblOption.Visible = False
    lblCurrent.Visible = False
End Sub

Private Sub cmdOK_Click()
' This subroutine sets the global variables to the new/selected
' label ormat and then closes the label format dialog box.

    ' Set the Label Specs to the new/selected label format
    '=======================
    ' All defined in globals
    strLabelFormat = Trim$(txtDescription.Text)
    sngTopMargin = Val(txtTop.Text)
    sngLeftMargin = Val(txtLeft.Text)
    sngWidth = Val(txtWidth.Text)
    sngHeight = Val(txtHeight.Text)
    sngSpacingTB = Val(txtSpacingTB.Text)
    sngSpacingRL = Val(txtSpacingRL.Text)
    intLines = Val(txtLines.Text)
    intLabelsPerRow = Val(txtLabelsPerRow.Text)
    intOptical = chkOptical.Value
    intAutoSize = chkAutoSize.Value
    '=======================
    
    ' Close the label format dialog box.
    Unload frmFormat
End Sub

Private Sub cmdOk_GotFocus()
' This subroutine hides the paste button, and the option and current labels.

    cmdPaste.Visible = False
    lblOption.Visible = False
    lblCurrent.Visible = False
End Sub

Private Sub cmdPaste_Click()
' This option pastes the current text of the paste button into
' the currently selected text box.

    txtSelected.Text = FormatLabelDimension(cmdPaste.Caption)
    txtSelected.SetFocus
End Sub

Private Sub Form_Activate()
' This subroutine places the current label format options
' into their respective text boxes when the form is shown.

    txtTop.Text = FormatLabelDimension(sngTopMargin)
    txtLeft.Text = FormatLabelDimension(sngLeftMargin)
    txtWidth.Text = FormatLabelDimension(sngWidth)
    txtHeight.Text = FormatLabelDimension(sngHeight)
    txtSpacingRL.Text = FormatLabelDimension(sngSpacingRL)
    txtSpacingTB.Text = FormatLabelDimension(sngSpacingTB)
    
    txtLabelsPerRow.Text = intLabelsPerRow
    txtLines.Text = intLines
    chkOptical.Value = intOptical
    txtDescription.Text = strLabelFormat
    chkAutoSize.Value = intAutoSize
    
    If blnLockFormat Then
        cmdOk.Enabled = False
        mnuDelete.Enabled = False
    Else
        cmdOk.Enabled = True
        mnuDelete.Enabled = True
    End If
    
    txtTop.SetFocus
    
    optWire.Value = True
    optWire.Enabled = False
    optTerminal.Enabled = False

End Sub

Private Sub Form_Load()
' This subroutine loads the formats that are saved into the
' list box and then finds the currently selected label format
' if it was saved before.

    On Error GoTo ErrorHandler
    
    Dim strTemp As String
    
    ' Open the label formats file for input to read in all
    ' of the label formats that have been saved.
    Open App.Path & "\" & LABEL_FORMATS For Input As #1
    If EOF(1) Then
        Close #1
        Exit Sub
    End If
    
    lstFormats.Clear
    
    ' Add each label format to the list.
    Do
        Input #1, strTemp
        ' Search for the "~" that precedes each label format name
        If Left$(strTemp, 1) = "~" Then
            lstFormats.AddItem Trim$(Mid$(strTemp, 2))
            If UCase$(Trim$(Mid$(strTemp, 2))) = UCase$(Trim$(strLabelFormat)) Then
                lstFormats.Selected(lstFormats.NewIndex) = True
            End If
        End If
    Loop Until EOF(1)
    Close #1
    
    ' Load the current label format settings into the text boxes.
    txtTop.Text = FormatLabelDimension(sngTopMargin)
    txtLeft.Text = FormatLabelDimension(sngLeftMargin)
    txtWidth.Text = FormatLabelDimension(sngWidth)
    txtHeight.Text = FormatLabelDimension(sngHeight)
    txtSpacingRL.Text = FormatLabelDimension(sngSpacingRL)
    txtSpacingTB.Text = FormatLabelDimension(sngSpacingTB)
    
    txtTLeft.Text = FormatLabelDimension(sngLeftMargin)
    txtTWidth.Text = FormatLabelDimension(sngWidth)
    txtTHeight.Text = FormatLabelDimension(sngHeight)
    txtTSpacing.Text = FormatLabelDimension(sngSpacingTB)

    
    txtLabelsPerRow.Text = intLabelsPerRow
    txtLines.Text = intLines
    txtDescription.Text = strLabelFormat
    If intOptical = True Then
        chkOptical.Value = 1
    Else
        chkOptical.Value = 0
    End If
    Exit Sub

ErrorHandler:
    If Err.Number = 53 Then
        Close #1
    Else
        MsgBox "Error " & Err.Number & " - " & Err.Description, , "Error"
    End If
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
' This subroutine closes the calculator if it is still running
' when we close out the label format dialog box in any way.

    If cmdCalc.Enabled = False Then
        Call TerminateProcess(lngProc, lngExit)
    End If
End Sub

Private Sub frmTerminal_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lstFormats_Click()

    ' If there are no formats then just exit the subroutine.
    If lstFormats.ListCount = 0 Then
        Exit Sub
    End If

    '200510 N.F.
    '   Abstracted the label spec load routine so that we now only store the
    '   format name in the registry, and load the actual values from label spec file
    '   on first run of program.
    On Error GoTo ErrorHandler

    Dim sTemp(10) As String
    
    If SelectNewFormat(Trim$(lstFormats.Text), sTemp) Then
        txtTop.Text = FormatLabelDimension(sTemp(0))
        txtLeft.Text = FormatLabelDimension(sTemp(1))
        txtWidth.Text = FormatLabelDimension(sTemp(2))
        txtHeight.Text = FormatLabelDimension(sTemp(3))
        txtSpacingTB.Text = FormatLabelDimension(sTemp(4))
        txtSpacingRL.Text = FormatLabelDimension(sTemp(5))
        txtLines.Text = Int(Val(sTemp(6)))
        txtLabelsPerRow.Text = Int(Val(sTemp(7)))
        chkOptical.Value = Val(sTemp(8))
        chkAutoSize.Value = Val(sTemp(9))
        ' Set the Description text box to the name of the label format.
        txtDescription.Text = Trim$(lstFormats.Text)
    Else
        Debug.Assert False
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & " - " & Err.Description, , "Error"
End Sub

Private Sub lstFormats_GotFocus()
' This subroutine hides the paste button and changes the
' currently selected option label to "Saved Label Formats".
    
    cmdPaste.Visible = False
    lblOption.Caption = "Saved Label Formats"
End Sub

Private Sub lstFormats_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub mnuDelete_Click()
' This subroutine allows the user to delete one of the formats
' that has been previously saved.

    On Error GoTo ErrorHandler
    
    Dim strTemp As String
    Dim intCnt As Long
    Dim Result As Long
    
    ' If the description of the label is "New Format" then the file has
    ' never actually been saved before.
    If txtDescription.Text = "New Format" Then
        MsgBox "You cannot delete a format that has not been saved!", vbOKOnly + vbInformation, "Cannot Delete!"
        Exit Sub
    End If
    
    ' Warn the user that they are about to delete a format and let them cancel if they choose to.
    Result = MsgBox("You are about to delete the format " _
             & Chr(34) & txtDescription.Text & Chr(34) & "!" _
             & Chr(13) & "Are you sure?", vbYesNo, "Delete Format?")
    
    ' If they choose not to delete the format, then just exit this subroutine.
    If Result = vbNo Then
        Exit Sub
    End If
    
    ' If there was only one format in the file, then just delete
    ' the label format file and clear the format list.
    If lstFormats.ListCount = 1 Then
        Kill App.Path & "\" & LABEL_FORMATS
        lstFormats.Clear
        Exit Sub
    End If
    
    ' Delete the temporary formats file if it exists.
    Kill App.Path & "\formats.tmp"
    
    ' Rename the current formats file to the temporary file name.
    Name App.Path & "\" & LABEL_FORMATS As App.Path & "\formats.tmp"
    
    ' Open the temporary file for input and a new formats file for output.
    Open App.Path & "\formats.tmp" For Input As #1
    Open App.Path & "\" & LABEL_FORMATS For Output As #2
    
    ' Loop through file, saving all the formats except the selected one.
    Do
        ' Read a line form the input file.
        Input #1, strTemp
        
        ' If the format matches the format we are deleting, then
        ' read the value lines following it.  Otherwise, write the
        ' line to the new formats file.
        If Trim$(strTemp) = "~" & txtDescription.Text Then
            For intCnt = 1 To 10
                Input #1, strTemp
            Next intCnt
        Else
            Print #2, Trim$(strTemp)
        End If
    Loop Until EOF(1)               ' Continue until we get to the end of the file.
    
    ' Close both of the input and output files.
    Close #1
    Close #2
    
    ' Delete the temporary file we created.
    Kill App.Path & "\formats.tmp"
    
    ' Remove the item we chose to delete from the list.
    lstFormats.RemoveItem lstFormats.ListIndex
    lstFormats.Selected(0) = True
    
    Exit Sub
ErrorHandler:
    If Err.Number = 53 Then
        Resume Next
    Else
        MsgBox "Error " & Err.Number & " - " & Err.Description, , "Error"
    End If
End Sub

Private Sub mnuSave_Click()
' This subroutine allows the user to save the settings that
' they currently have on the screen to the formats file.

    On Error GoTo ErrorHandler
    
    Dim strFormat As String
    Dim strTemp As String
    Dim intCnt As Long
    Dim intTemp As Long
    Dim intResult As Long
    
    ' Format the text so it is ready to write to the file.
    txtTop.Text = FormatLabelDimension(txtTop.Text)
    txtLeft.Text = FormatLabelDimension(txtLeft.Text)
    txtWidth.Text = FormatLabelDimension(txtWidth.Text)
    txtHeight.Text = FormatLabelDimension(txtHeight.Text)
    txtSpacingRL.Text = FormatLabelDimension(txtSpacingRL.Text)
    txtSpacingTB.Text = FormatLabelDimension(txtSpacingTB.Text)
    
    txtLabelsPerRow.Text = Int(Val(txtLabelsPerRow.Text))
    txtLines.Text = Int(Val(txtLines.Text))
    
    ' Prompt the user to input a name for the label format.
    strFormat = InputBox("Enter New Format Name:", "Save Format", Trim$(txtDescription.Text))
    
    ' If they didn't enter anything or clicked cancel then just exit out of the subroutine.
    If Trim$(strFormat) = "" Then
        Exit Sub
    End If
    
    ' Check the current label format list to see if the format already exists.
    ' If it does, just let the user know that they have to choose a different name.
    For intCnt = 0 To lstFormats.ListCount - 1
        If UCase$(Trim$(lstFormats.List(intCnt))) = UCase$(Trim$(strFormat)) Then
            intResult = MsgBox("Format already exists!" & vbCrLf & "Do you wish to overwrite?", vbYesNo + vbInformation, "Already Exists")
            If intResult = vbNo Then
                Exit Sub
            Else
                On Error Resume Next
                ' If there was only one format in the file, then just delete
                ' the label format file and clear the format list.
                If lstFormats.ListCount = 1 Then
                    Kill App.Path & "\" & LABEL_FORMATS
                    lstFormats.Clear
                    Exit Sub
                End If
                
                ' Delete the temporary formats file if it exists.
                Kill App.Path & "\formats.tmp"
                
                ' Rename the current formats file to the temporary file name.
                Name App.Path & "\" & LABEL_FORMATS As App.Path & "\formats.tmp"
                
                ' Open the temporary file for input and a new formats file for output.
                Open App.Path & "\formats.tmp" For Input As #1
                Open App.Path & "\" & LABEL_FORMATS For Output As #2
                
                ' Loop through file, saving all the formats except the selected one.
                Do
                    ' Read a line from the input file.
                    Input #1, strTemp
                    
                    ' If the format matches the format we are deleting, then
                    ' read the value lines following it.  Otherwise, write the
                    ' line to the new formats file.
                    If Trim$(strTemp) = "~" & Trim$(strFormat) Then
                        For intTemp = 1 To 10
                            Input #1, strTemp
                        Next intTemp
                    Else
                        Print #2, Trim$(strTemp)
                    End If
                Loop Until EOF(1)               ' Continue until we get to the end of the file.
                
                ' Close both of the input and output files.
                Close #1
                Close #2
                
                ' Delete the temporary file we created.
                Kill App.Path & "\formats.tmp"
            End If
            On Error GoTo ErrorHandler
        End If
    Next intCnt
    
    ' If the name of the format is still "New Format" then
    ' tell the user they must select a different name, because
    ' that is the one the program uses to tell if a format has
    ' been saved yet or not.
    If UCase$(Trim$(strFormat)) = "NEW FORMAT" Then
        MsgBox "Please Select Different Format Name"
        mnuSave_Click
        Exit Sub
    End If
    
    ' Open the format file for appending output.
    Open App.Path & "\" & LABEL_FORMATS For Append As #1
    
    ' Write the new format to the file.
    Print #1, "~" & Trim$(strFormat)
    Print #1, Trim$(txtTop.Text)
    Print #1, Trim$(txtLeft.Text)
    Print #1, Trim$(txtWidth.Text)
    Print #1, Trim$(txtHeight.Text)
    Print #1, Trim$(txtSpacingTB.Text)
    Print #1, Trim$(txtSpacingRL.Text)
    Print #1, Trim$(txtLines.Text)
    Print #1, Trim$(txtLabelsPerRow.Text)
    Print #1, Trim$(Format(chkOptical.Value, "0"))
    Print #1, Trim$(Format(chkAutoSize.Value, "0"))
    
    ' Close the file we just wrote the settings to.
    Close #1
    
    If intResult <> vbYes Then
        ' Add the current label format to the list box.
        lstFormats.AddItem Trim$(strFormat)
        lstFormats.Selected(lstFormats.NewIndex) = True
        lstFormats_Click
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & " - " & Err.Description, , "Error"
End Sub

Private Sub optTerminal_Click()
    fraTerminal.ZOrder (1)
    fraWire.ZOrder (1)
    fraOptions.ZOrder (0)
    lblTOption.Caption = ""
End Sub

Private Sub optWire_Click()
    fraWire.ZOrder (1)
    fraTerminal.ZOrder (1)
    fraOptions.ZOrder (0)
    lblOption.Caption = ""
End Sub



Private Sub timCalc_Timer()
' This subroutine constantly checks to see if the calculator program
' is still running.  If it is, it reads the new value out of the
' calculator's text box and enters it into the paste button caption.
' If the user closes the calculator, then the program re-enables the
' calculator button and hides the paste button.

    Dim strCalc As String
    Dim rectCalc As RECT
    Dim lngTemp As Long
    
    DoEvents
    
    ' Find any window that has the title "Calculator"
    lngTemp = FindWindow(vbNullString, "Calculator")
    
    ' If there is one titled "Calculator" and there aren't any
    ' titled "Kasa Wire Labels Calculator" then change the window
    ' text of the "Calculator" window to "Kasa Wire Labels Calculator"
    ' and re-obtain the settings of that window.
    If lngTemp <> 0 And FindWindow(vbNullString, "Kasa Wire Labels Calculator") = 0 Then
        ' Change the text of the calculator window to "Kasa Wire Labels Calculator"
        Call SetWindowText(lngTemp, "Kasa Wire Labels Calculator")
    
        ' Get the rectangular coordinates of the calculator program
        Call GetWindowRect(lngTemp, rectCalc)
    
        ' Set the calculator to always on top
        Call SetWindowPos(lngTemp, -1, rectCalc.Left, rectCalc.Top, (rectCalc.Right - rectCalc.Left), (rectCalc.Bottom - rectCalc.Top), 0)
    
        ' Find the handle to the calculator's text box
        '200510 N.F.
        'A value of y = 66 covers all scenarios except XP extra large
        lngWindow = WindowFromPoint(rectCalc.Left + 232, rectCalc.Top + 66)
        If lngWindow = lngCalc Then
            lngWindow = WindowFromPoint(rectCalc.Left + 232, rectCalc.Top + 75)
        End If
        
        lngCalc = lngTemp
    End If
    
    ' There is no calculator program running, so enable the calculator
    ' button and hide/disable the paste button.  Plus, quit checking
    ' for a value in a calculator program.
    If FindWindow(vbNullString, "Kasa Wire Labels Calculator") = 0 Then
        timCalc.Enabled = False
        cmdPaste.Enabled = False
        cmdPaste.Visible = False
        cmdCalc.Enabled = True
        Exit Sub
    End If
    
    ' Get the length of the text in the calculator text box.
    lngLength = SendMessage(lngWindow, WM_GETTEXTLENGTH, ByVal CLng(0), ByVal CLng(0)) + 1
    strCalc = Space$(lngLength)
    
    ' Retrieve the text of the calculator text box.
    SendMessage lngWindow, WM_GETTEXT, ByVal lngLength, ByVal strCalc
    
    lngLength = Len(strCalc) - 1
    strCalc = Left$(Trim$(strCalc), lngLength)
    
    ' Change the text in the paste button if the text is different.
    If Val(cmdPaste.Caption) <> Val(Left$(strCalc, 7)) Then
        cmdPaste.Caption = Val(Left$(strCalc, 7))
    End If
End Sub

Private Sub txtDescription_Click()
' This subroutine selects all the text in the text box.
    SelectText
End Sub

Private Sub txtDescription_GotFocus()
' This subroutine hides the paste button, selects all the text
' in the text box, and changes the option caption to "Label Format Description".

    cmdPaste.Visible = False
    lblOption.Caption = "Label Format Description"
    SelectText
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtDescription_LostFocus()
' This subroutine cleans up the text in the description text box.

    txtDescription.Text = Trim$(txtDescription.Text)
End Sub

Private Sub txtHeight_Change()
' This subroutine changes the label format to "New Format" when
' the text in txtHeight changes.

    txtDescription.Text = "New Format"
End Sub

Private Sub txtHeight_GotFocus()
' This subroutine changes the selected option label to
' "Label Height" and sets the Selected Text Box to txtHeight
' for the calculator to work.  It also highlights the arrows
' associated with the Label Height

    Dim intTemp As Long
    
    lblOption.Visible = True
    lblCurrent.Visible = True
    lblOption.Caption = "Label Height"
    
    ' Set this text box to the Selected text box
    Set txtSelected = txtHeight
    
    ' Move the paste button next to the height text box.
    With cmdPaste
        If timCalc.Enabled Then
            .Visible = True
        End If
        .Left = txtHeight.Left + H_OFFSET
        .Top = txtHeight.Top + txtHeight.Height + V_OFFSET
        .Width = txtHeight.Width
    End With
    
    ' Highlight the text in the text box
    SelectText
    
    ' Highlight the arrows associated with the label height.
    For intTemp = 0 To 6
        linHeight(intTemp).BorderColor = vbRed
        linHeight(intTemp).BorderWidth = 2
    Next intTemp
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
    And KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight _
    And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab _
    And KeyAscii <> vbKeyDecimal Then
        KeyAscii = 0
    ElseIf KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab Then
        txtDescription.Text = "New Format"
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtHeight_LostFocus()
' This subroutine puts the txtHeight value to the appropriate
' format and the un-highlights the arrows associated with the
' label height.

    Dim intTemp As Long
    
    ' Format the text correctly
    txtHeight.Text = FormatLabelDimension(txtHeight.Text)
    
    ' Un-highlight the arrows associated with the label height.
    For intTemp = 0 To 6
        linHeight(intTemp).BorderColor = vbBlack
        linHeight(intTemp).BorderWidth = 1
    Next intTemp
End Sub

Private Sub txtLabelsPerRow_Change()
' This subroutine sets the label description to "New Format"
' if the value in LabelsPerRow changes.

    txtDescription.Text = "New Format"
End Sub

Private Sub txtLabelsPerRow_GotFocus()
' This subroutine changes the selected option label to
' "Labels Per Row" and highlights the text

    cmdPaste.Visible = False
    lblOption.Caption = "Labels Per Row"
    SelectText
End Sub

Private Sub txtLabelsPerRow_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
    And KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight _
    And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab _
    And KeyAscii <> vbKeyDecimal Then
        KeyAscii = 0
    ElseIf KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab Then
        txtDescription.Text = "New Format"
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtLabelsPerRow_LostFocus()
' This subroutine puts the value in the text box to the
' appropriate format.

    txtLabelsPerRow.Text = Int(Val(txtLabelsPerRow.Text))
    If Val(txtLabelsPerRow.Text) <= 0 Then
        txtLabelsPerRow.Text = "1"
    End If
End Sub

Private Sub txtLeft_Change()
' This subroutine sets the label description to "New Format"
' if the text in txtLeft changes.

    txtDescription.Text = "New Format"
End Sub

Private Sub txtLeft_GotFocus()
' This subroutine changes the selected option label to
' "Left Margin" and sets the Selected Text Box to txtLeft
' for the calculator to work.  It also highlights the arrows
' associated with the Left Margin.
    
    Dim intTemp As Long
    
    lblOption.Visible = True
    lblCurrent.Visible = True
    lblOption.Caption = "Left Margin"
    
    ' Set the selected text box to txtLeft
    Set txtSelected = txtLeft
    
    ' Move the paste button next to the left margin text box.
    With cmdPaste
        If timCalc.Enabled Then
            .Visible = True
        End If
        .Left = txtLeft.Left + H_OFFSET
        .Top = txtLeft.Top + txtLeft.Height + V_OFFSET
        .Width = txtLeft.Width
    End With
    
    ' Highlight all the text
    SelectText
    
    ' Highlight the arrows
    For intTemp = 0 To 6
        linLeftMar(intTemp).BorderColor = vbRed
        linLeftMar(intTemp).BorderWidth = 2
    Next intTemp
End Sub

Private Sub txtLeft_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
    And KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight _
    And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab _
    And KeyAscii <> vbKeyDecimal Then
        KeyAscii = 0
    ElseIf KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab Then
        txtDescription.Text = "New Format"
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtLeft_LostFocus()
' This subroutine formats the left format text to the correct
' format and un-highlights the arrows.

    Dim intTemp As Long
    
    ' Format the text correctly
    txtLeft.Text = FormatLabelDimension(txtLeft.Text)
    
    ' Un-Hightlight the arrows.
    For intTemp = 0 To 6
        linLeftMar(intTemp).BorderColor = vbBlack
        linLeftMar(intTemp).BorderWidth = 1
    Next intTemp
End Sub

Private Sub txtLines_Change()
' This subroutine changes the label description to "New Format"
' if the text in txtLines changes.

    txtDescription.Text = "New Format"
End Sub

Private Sub txtLines_GotFocus()
' This subroutine changes the selected option label to
' "Printed Lines Per Label" and selects all the text

    cmdPaste.Visible = False
    lblOption.Caption = "Printed Lines Per Label"
    SelectText
End Sub

Private Sub txtLines_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
    And KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight _
    And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab _
    And KeyAscii <> vbKeyDecimal Then
        KeyAscii = 0
    ElseIf KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab Then
        txtDescription.Text = "New Format"
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtLines_LostFocus()
' This subroutine formats the text to the correct format.

    txtLines.Text = Int(Val(txtLines.Text))
    If Val(txtLines.Text) <= 0 Then
        txtLines.Text = "1"
    End If
End Sub

Private Sub txtSpacingRL_Change()
' This subroutine changes the label description to "New Format"
' if the text in txtSpacingRL changes.

    txtDescription.Text = "New Format"
End Sub

Private Sub txtSpacingRL_GotFocus()
' This subroutine changes the selected option label to
' "Label Spacing - Right to Left" and sets the Selected Text Box to txtSpacingRL
' for the calculator to work.  It also highlights the arrows
' associated with the Horizontal Spacing.
    
    Dim intTemp As Long
    
    lblOption.Visible = True
    lblCurrent.Visible = True
    lblOption.Caption = "Label Spacing - Right to Left"
    
    ' Set the selected text box to txtSpacingRL
    Set txtSelected = txtSpacingRL
    
    ' Move the paste button next to txtSpacingRL
    With cmdPaste
        If timCalc.Enabled Then
            .Visible = True
        End If
        .Left = txtSpacingRL.Left + H_OFFSET
        .Top = txtSpacingRL.Top + txtSpacingRL.Height + V_OFFSET
        .Width = txtSpacingRL.Width
    End With
    
    ' Select all the text
    SelectText
    
    ' Hightlight the arrows
    For intTemp = 0 To 6
        linSpacingRL(intTemp).BorderColor = vbRed
        linSpacingRL(intTemp).BorderWidth = 2
    Next intTemp
End Sub

Private Sub txtSpacingRL_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
    And KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight _
    And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab _
    And KeyAscii <> vbKeyDecimal Then
        KeyAscii = 0
    ElseIf KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab Then
        txtDescription.Text = "New Format"
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtSpacingRL_LostFocus()
' This subroutine formats the text to the correct format,
' and then un-highlights the arrows.

    Dim intTemp As Long
    
    txtSpacingRL.Text = FormatLabelDimension(txtSpacingRL.Text)

    For intTemp = 0 To 6
        linSpacingRL(intTemp).BorderColor = vbBlack
        linSpacingRL(intTemp).BorderWidth = 1
    Next intTemp
End Sub

Private Sub txtSpacingTB_Change()
' This subroutine changes the label description to "New Format"
' when the value of SpacingTB changes.

    txtDescription.Text = "New Format"
End Sub

Private Sub txtSpacingTB_GotFocus()
' This subroutine sets the selected text box to txtSpacingTB
' and moves the paste button next to the text box.  It also
' selects all the text and highlights the arrows.

    Dim intTemp As Long
    
    lblOption.Visible = True
    lblCurrent.Visible = True
    lblOption.Caption = "Label Spacing - Top to Bottom"
    
    ' Set the selected text box to txtSpacingTB
    Set txtSelected = txtSpacingTB
    
    ' Move the paste command button
    With cmdPaste
        If timCalc.Enabled Then
            .Visible = True
        End If
        .Left = txtSpacingTB.Left + H_OFFSET
        .Top = txtSpacingTB.Top + txtSpacingTB.Height + V_OFFSET
        .Width = txtSpacingTB.Width
    End With
    
    ' Select all the text
    SelectText
    
    ' Highlight the arrows
    For intTemp = 0 To 6
        linSpacingTB(intTemp).BorderColor = vbRed
        linSpacingTB(intTemp).BorderWidth = 2
    Next intTemp
End Sub

Private Sub txtSpacingTB_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
    And KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight _
    And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab _
    And KeyAscii <> vbKeyDecimal Then
        KeyAscii = 0
    ElseIf KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab Then
        txtDescription.Text = "New Format"
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtSpacingTB_LostFocus()
' This subroutine formats the text to the correct format and
' un-highlights the arrows.

    Dim intTemp As Long
    
    txtSpacingTB.Text = FormatLabelDimension(txtSpacingTB.Text)

    For intTemp = 0 To 6
        linSpacingTB(intTemp).BorderColor = vbBlack
        linSpacingTB(intTemp).BorderWidth = 1
    Next intTemp
End Sub

Private Sub txtTHeight_Change()
    txtDescription.Text = "New Format"
End Sub

Private Sub txtTHeight_GotFocus()
' This subroutine changes the selected option label to
' "Label Height" and sets the Selected Text Box to txtHeight
' for the calculator to work.  It also highlights the arrows
' associated with the Label Height

    Dim intTemp As Long
    
    lblOption.Visible = True
    lblCurrent.Visible = True
    lblTOption.Caption = "Label Height"
    
    ' Set this text box to the Selected text box
    Set txtSelected = txtTHeight
    
    ' Move the paste button next to the height text box.
    With cmdPaste
        If timCalc.Enabled Then
            .Visible = True
        End If
        .Left = txtTHeight.Left + H_OFFSET + TH_OFFSET
        .Top = txtTHeight.Top + txtTHeight.Height + V_OFFSET + TV_OFFSET
        .Width = txtTHeight.Width
    End With
    
    ' Highlight the text in the text box
    SelectText
    
    ' Highlight the arrows associated with the label height.
    For intTemp = 0 To 7
        lneHeight(intTemp).BorderColor = vbRed
        lneHeight(intTemp).BorderWidth = 2
    Next intTemp
End Sub


Private Sub txtTHeight_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
    And KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight _
    And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab _
    And KeyAscii <> vbKeyDecimal Then
        KeyAscii = 0
    ElseIf KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab Then
        txtDescription.Text = "New Format"
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtTHeight_LostFocus()
' This subroutine puts the txtHeight value to the appropriate
' format and the un-highlights the arrows associated with the
' label height.

    Dim intTemp As Long
    
    ' Format the text correctly
    txtTHeight.Text = FormatLabelDimension(txtTHeight.Text)
    
    ' Un-highlight the arrows associated with the label height.
    For intTemp = 0 To 7
        lneHeight(intTemp).BorderColor = vbBlack
        lneHeight(intTemp).BorderWidth = 1
    Next intTemp
End Sub

Private Sub txtTLeft_Change()
    txtDescription.Text = "New Format"
End Sub

Private Sub txtTLeft_GotFocus()
' This subroutine changes the selected option label to
' "Left Margin" and sets the Selected Text Box to txtLeft
' for the calculator to work.  It also highlights the arrows
' associated with the Left Margin.
    
    Dim intTemp As Long
    
    'lblOption.Visible = True
    'lblCurrent.Visible = True
    lblTOption.Caption = "Left Margin"
    
    ' Set the selected text box to txtLeft
    Set txtSelected = txtTLeft
    
    ' Move the paste button next to the left margin text box.
    With cmdPaste
        If timCalc.Enabled Then
            .Visible = True
        End If
        .Left = txtTLeft.Left + H_OFFSET + TH_OFFSET
        .Top = txtTLeft.Top + txtTLeft.Height + V_OFFSET + TV_OFFSET
        .Width = txtTLeft.Width
    End With
    
    ' Highlight all the text
    SelectText
    
    ' Highlight the arrows
    For intTemp = 0 To 4
        lneLeftMar(intTemp).BorderColor = vbRed
        lneLeftMar(intTemp).BorderWidth = 2
    Next intTemp
End Sub

Private Sub txtTLeft_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
    And KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight _
    And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab _
    And KeyAscii <> vbKeyDecimal Then
        KeyAscii = 0
    ElseIf KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab Then
        txtDescription.Text = "New Format"
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtTLeft_LostFocus()
' This subroutine formats the left format text to the correct
' format and un-highlights the arrows.

    Dim intTemp As Long
    
    ' Format the text correctly
    txtTLeft.Text = FormatLabelDimension(txtTLeft.Text)
    
    ' Un-Hightlight the arrows.
    For intTemp = 0 To 4
        lneLeftMar(intTemp).BorderColor = vbBlack
        lneLeftMar(intTemp).BorderWidth = 1
    Next intTemp
End Sub

Private Sub txtTop_Change()
' This subroutine changes the label description to "New Format"
' when the value of the top margin changes.

    txtDescription.Text = "New Format"
End Sub

Private Sub txtTop_GotFocus()
' This subroutine sets the selected text box to txtTop and
' moves the paste button next to the text box.  It also selects
' all the text in the text box and highlights the arrows.

    Dim intTemp As Long
    
    lblOption.Visible = True
    lblCurrent.Visible = True
    lblOption.Caption = "Top Margin"
    
    ' Set the selected text to txtTop
    Set txtSelected = txtTop
    
    ' Move the paste button next to the text box
    With cmdPaste
        If timCalc.Enabled Then
            .Visible = True
        End If
        .Left = txtTop.Left + txtTop.Width + H_OFFSET
        .Top = txtTop.Top + V_OFFSET
        .Height = txtTop.Height
    End With
    
    ' Select all the text
    SelectText
    
    ' Highlight the arrows
    For intTemp = 0 To 6
        linTopMar(intTemp).BorderColor = vbRed
        linTopMar(intTemp).BorderWidth = 2
    Next intTemp
End Sub

Private Sub txtTop_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
    And KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight _
    And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab _
    And KeyAscii <> vbKeyDecimal Then
        KeyAscii = 0
    ElseIf KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab Then
        txtDescription.Text = "New Format"
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtTop_LostFocus()
' This subroutine formats the text to the correct format.
' Then it un-highlights the arrows.

    Dim intTemp As Long
    
    txtTop.Text = FormatLabelDimension(txtTop.Text)
    For intTemp = 0 To 6
        linTopMar(intTemp).BorderColor = vbBlack
        linTopMar(intTemp).BorderWidth = 1
    Next intTemp
End Sub

Private Sub txtTSpacing_Change()
        txtDescription.Text = "New Format"
End Sub

Private Sub txtTSpacing_GotFocus()
' This subroutine sets the selected text box to txtSpacingTB
' and moves the paste button next to the text box.  It also
' selects all the text and highlights the arrows.

    Dim intTemp As Long
    
    'lblOption.Visible = True
    'lblCurrent.Visible = True
    lblTOption.Caption = "Label Spacing - Center to Center"
    
    ' Set the selected text box to txtSpacingTB
    Set txtSelected = txtTSpacing
    
    ' Move the paste command button
    With cmdPaste
        If timCalc.Enabled Then
            .Visible = True
        End If
        .Left = txtTSpacing.Left + H_OFFSET + TH_OFFSET
        .Top = txtTSpacing.Top + txtSpacingTB.Height + V_OFFSET + TV_OFFSET
        .Width = txtTSpacing.Width
    End With
    
    ' Select all the text
    SelectText
    
    ' Highlight the arrows
    For intTemp = 0 To 7
        lneSpacing(intTemp).BorderColor = vbRed
        lneSpacing(intTemp).BorderWidth = 2
    Next intTemp
End Sub

Private Sub txtTSpacing_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
    And KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight _
    And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab _
    And KeyAscii <> vbKeyDecimal Then
        KeyAscii = 0
    ElseIf KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab Then
        txtDescription.Text = "New Format"
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtTSpacing_LostFocus()
' This subroutine formats the text to the correct format and
' un-highlights the arrows.

    Dim intTemp As Long
    
    txtTSpacing.Text = FormatLabelDimension(txtTSpacing.Text)

    For intTemp = 0 To 7
        lneSpacing(intTemp).BorderColor = vbBlack
        lneSpacing(intTemp).BorderWidth = 1
    Next intTemp
End Sub

Private Sub txtTWidth_Change()
    txtDescription.Text = "New Format"
End Sub

Private Sub txtTWidth_GotFocus()
' This subroutine sets the selected text box to txtWidth and
' moves the paste button next to the text box.  It also selects
' all the text and highlights the arrows.

    Dim intTemp As Long
    
    lblOption.Visible = True
    lblCurrent.Visible = True
    lblTOption.Caption = "Label Width"
    
    ' Set the selected text box to txtWidth
    Set txtSelected = txtTWidth
    
    ' Move the paste button next to the text box
    With cmdPaste
        If timCalc.Enabled Then
            .Visible = True
        End If
        .Left = txtTWidth.Left + H_OFFSET + TH_OFFSET
        .Top = txtTWidth.Top + txtTWidth.Height + V_OFFSET + TV_OFFSET
        .Height = txtTWidth.Height
    End With
    
    ' Select all the text
    SelectText
    
    ' Highlight the arrows
    For intTemp = 0 To 8
        lneWidth(intTemp).BorderColor = vbRed
        lneWidth(intTemp).BorderWidth = 2
    Next intTemp
End Sub

Private Sub txtTWidth_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
    And KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight _
    And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab _
    And KeyAscii <> vbKeyDecimal Then
        KeyAscii = 0
    ElseIf KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab Then
        txtDescription.Text = "New Format"
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtTWidth_LostFocus()
' This subroutine formats the text to the correct format and then
' un-highlights the arrows.

    Dim intTemp As Long
    
    txtTWidth.Text = FormatLabelDimension(txtTWidth.Text)
    For intTemp = 0 To 8
        lneWidth(intTemp).BorderColor = vbBlack
        lneWidth(intTemp).BorderWidth = 1
    Next intTemp
End Sub

Private Sub txtWidth_Change()
' This subroutine sets the label description to "New Format"
' whenever the value of txtWidth changes.

    txtDescription.Text = "New Format"
End Sub

Private Sub txtWidth_GotFocus()
' This subroutine sets the selected text box to txtWidth and
' moves the paste button next to the text box.  It also selects
' all the text and highlights the arrows.

    Dim intTemp As Long
    
    lblOption.Visible = True
    lblCurrent.Visible = True
    lblOption.Caption = "Label Width"
    
    ' Set the selected text box to txtWidth
    Set txtSelected = txtWidth
    
    ' Move the paste button next to the text box
    With cmdPaste
        If timCalc.Enabled Then
            .Visible = True
        End If
        .Left = txtWidth.Left + txtWidth.Width + H_OFFSET
        .Top = txtWidth.Top + V_OFFSET
        .Height = txtWidth.Height
    End With
    
    ' Select all the text
    SelectText
    
    ' Highlight the arrows
    For intTemp = 0 To 6
        linWidth(intTemp).BorderColor = vbRed
        linWidth(intTemp).BorderWidth = 2
    Next intTemp
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
    And KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight _
    And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab _
    And KeyAscii <> vbKeyDecimal Then
        KeyAscii = 0
    ElseIf KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab Then
        txtDescription.Text = "New Format"
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
    
End Sub


Private Sub txtWidth_LostFocus()
' This subroutine formats the text to the correct format and then
' un-highlights the arrows.

    Dim intTemp As Long
    
    txtWidth.Text = FormatLabelDimension(txtWidth.Text)
    For intTemp = 0 To 6
        linWidth(intTemp).BorderColor = vbBlack
        linWidth(intTemp).BorderWidth = 1
    Next intTemp
End Sub

Private Function GetWinPath()
' This function returns the full windows path.

    Dim strFolder As String
    Dim lngResult As Long

    strFolder = String$(MAX_PATH, 0)
    lngResult = GetWindowsDirectory(strFolder, MAX_PATH)

    If lngResult <> 0 Then
        GetWinPath = Left$(strFolder, InStr(strFolder, Chr(0)) - 1)
    Else
        GetWinPath = ""
    End If
End Function



