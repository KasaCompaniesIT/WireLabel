VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "KASA Wire Labels"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   -1020
   ClientWidth     =   8325
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   ScaleHeight     =   6510
   ScaleWidth      =   8325
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin MSComctlLib.Toolbar tlbToolbar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   28
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "Create New Label List"
            Object.ToolTipText     =   "Create New Label List"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open Label List"
            Object.ToolTipText     =   "Open Label List"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Import"
            Description     =   "Import Text File"
            Object.ToolTipText     =   "Import Text File"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Export"
            Description     =   "Export Text File"
            Object.ToolTipText     =   "Export Text File"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save Label File"
            Object.ToolTipText     =   "Save Label File"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Format"
            Description     =   "Change Label Format"
            Object.ToolTipText     =   "Change Label Format"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Description     =   "Print Labels"
            Object.ToolTipText     =   "Print Labels"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteLabel"
            Description     =   "Delete Label"
            Object.ToolTipText     =   "Delete Label"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Cut"
            Description     =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Copy"
            Description     =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Paste"
            Description     =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            Description     =   "Change Font"
            Object.ToolTipText     =   "Change Font"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Description     =   "Find Label"
            Object.ToolTipText     =   "Find Label"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sort"
            Description     =   "Sort Labels"
            Object.ToolTipText     =   "Sort Labels"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InsertChar"
            Description     =   "Insert Character"
            Object.ToolTipText     =   "Insert Character"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DeleteChar"
            Description     =   "Delete Character"
            Object.ToolTipText     =   "Delete Character"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Quantity"
            Description     =   "Change Quantity"
            Object.ToolTipText     =   "Change Quantity"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Decimal"
            Description     =   "Create Decimal Sequence"
            Object.ToolTipText     =   "Create Decimal Sequence"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Alpha"
            Description     =   "Create Alphabetic Sequence"
            Object.ToolTipText     =   "Create Alphabetic Sequence"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PLC"
            Description     =   "Create PLC Sequence"
            Object.ToolTipText     =   "Create PLC Sequence"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SLC"
            Description     =   "Create SLC Sequence"
            Object.ToolTipText     =   "Create SLC Sequence"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Description     =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   22
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   6270
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2822
            MinWidth        =   1764
            Text            =   "Label: "
            TextSave        =   "Label: "
            Key             =   "LabelNum"
            Object.ToolTipText     =   "Label Number (Current / Total)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3598
            Text            =   "Label Text:"
            TextSave        =   "Label Text:"
            Key             =   "LabelText"
            Object.ToolTipText     =   "Label Text"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2293
            MinWidth        =   1235
            Text            =   "Qty: "
            TextSave        =   "Qty: "
            Key             =   "LabelQty"
            Object.ToolTipText     =   "Label Quantity"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3527
            MinWidth        =   2469
            Text            =   "Total Qty: "
            TextSave        =   "Total Qty: "
            Key             =   "TotalQty"
            Object.ToolTipText     =   "Total Quantity of Labels"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   820
            MinWidth        =   820
            Text            =   "NUM"
            TextSave        =   "NUM"
            Key             =   "NumLock"
            Object.ToolTipText     =   "Number Lock (On/Off)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   900
            Text            =   "CAPS"
            TextSave        =   "CAPS"
            Key             =   "CapsLock"
            Object.ToolTipText     =   "Caps Lock (On/Off)"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLabelQty 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtLabelText 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cdlDialog 
      Left            =   5160
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   2
      FontName        =   "Arial"
   End
   Begin MSComctlLib.ProgressBar prgProgress 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   5
      Top             =   6045
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   5760
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1024
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1180
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1294
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C54
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D68
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2020
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":217C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2434
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2590
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2800
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvLabels 
      Height          =   6255
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11033
      View            =   3
      SortOrder       =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imlToolbar"
      SmallIcons      =   "imlToolbar"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "LabelText"
         Text            =   "Label Text"
         Object.Width           =   3309
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "LabelQty"
         Text            =   "Quantity"
         Object.Width           =   1059
      EndProperty
   End
   Begin VB.Label lblLabelQty 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblLabelText 
      Caption         =   "Label Text"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblUserMsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   480
      Width           =   5535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import Text File..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export Text File..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save..."
         Shortcut        =   ^S
      End
      Begin VB.Menu filMakeFieldTags 
         Caption         =   "&Make Field Tags..."
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuPrintSetup 
         Caption         =   "P&rinter Setup..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPrintPreview 
         Caption         =   "Print Pre&view..."
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuInsertBlank 
         Caption         =   "&Insert Blank Label"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChgQty 
         Caption         =   "Change Q&uantity..."
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuChgAllQty 
         Caption         =   "Change All &Quantities..."
      End
      Begin VB.Menu mnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertChar 
         Caption         =   "Insert C&haracters..."
      End
      Begin VB.Menu mnuDeleteChar 
         Caption         =   "Delete Characters..."
      End
      Begin VB.Menu mnuLine6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSort 
         Caption         =   "&Sort"
      End
      Begin VB.Menu mnuLabelFormat 
         Caption         =   "&Label Format..."
      End
      Begin VB.Menu mnuFont 
         Caption         =   "F&ont..."
      End
      Begin VB.Menu mnuDefaultQty 
         Caption         =   "Set Default Quantity..."
      End
   End
   Begin VB.Menu mnuSequence 
      Caption         =   "&Sequence"
      Begin VB.Menu mnuDecimal 
         Caption         =   "Decimal..."
      End
      Begin VB.Menu mnuAlpha 
         Caption         =   "&Alphabetic..."
      End
      Begin VB.Menu mnuPLC 
         Caption         =   "&PLC 5..."
      End
      Begin VB.Menu mnuSLC 
         Caption         =   "S&LC 500..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuLine8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolbar 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu mnuStatusBar 
         Caption         =   "&Status Bar"
      End
      Begin VB.Menu mnuProgressBar 
         Caption         =   "&Progress Bar"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupInsert 
         Caption         =   "&Insert Blank Label"
      End
      Begin VB.Menu mnuPopupCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuPopupCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPopupPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuPopupDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuPopupSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuPopupLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupInsertchar 
         Caption         =   "Insert C&haracters..."
      End
      Begin VB.Menu mnuPopupDeleteChar 
         Caption         =   "Delete Characters..."
      End
      Begin VB.Menu mnuPopupEditLabel 
         Caption         =   "&Edit Label"
      End
      Begin VB.Menu mnuPopupChgQty 
         Caption         =   "Change &Quantity..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'DBG NOTE: Set DebugMode in conditional compilation arguments.
'   DebugMode = 1 :  For displaying print preview for terminal strips (printstrips) routine.
'   DebugMode = 2 :  For debugging the mnuPrint routines.
#If DebugMode = 2 Then
    Dim TempText As String
    Dim TempText2 As String
#End If

'Type for WireLabels.
Private Type MyLabel
    'Holds 3 lines worth of text per Label.
    strLabel(2) As String
    'CHG 200508 NF : MaxText holds the maximum length text of strLabel() array.
    MaxText As String
    'Holds print font size for that label.  Size is determined by sizing to MaxText.
    lngSize As Long
End Type

'Declaration for adding lbl files that are opened to the Windows Recent Documents menu.
Private Const SHARD_PATH = 2
Private Declare Sub SHAddToRecentDocs Lib "shell32.dll" (ByVal uFlags As Long, ByVal PV As String)

' Declaration for having the computer wait a specified amount of time.
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Declarations for un-clicking the mouse button on the form.
Private Const WM_LBUTTONUP = &H202

'Printer object local to this form.
Private ptrDefault As Printer

' Module level variables.
Dim strFind As String                   ' Search String
Dim blnAllowPaste As Boolean            ' Allow Paste Option
Dim strOldLabel As String               ' Label Before Edit
Dim strFileName As String               ' Current Filename

Public Sub TotalQty()
' This subroutine updates the status bar at the bottom of the
' window, and it also updates the caption of the window with
' the application name and the name of the current file.

    Dim intCnt As Long
    Dim intTtlQty As Long
    
    intTtlQty = 0
    
    ' Count total quantity of labels
    For intCnt = 1 To lsvLabels.ListItems.Count
        intTtlQty = intTtlQty + Val(lsvLabels.ListItems.Item(intCnt).SubItems(1))
    Next intCnt
        
    ' Set window's caption to the app title and the current file
    Me.Caption = GetFormCaption(Me.Name) & " - " & strFileName
    stbStatus.Panels.Item(4).Text = "Total Qty: " & Trim$(Str$(intTtlQty))
End Sub

Public Sub CheckVisualStatus()
' This subroutine will update the menus that have
' check marks next to them to be consistent with the
' way that the window is displayed.  It also updates
' the title of the window with the current file name
' and makes sure that all controls are within the ranges
' that are predetermined in the Form_Resize subroutine.

    mnuProgressBar.Checked = prgProgress.Visible    ' If Progress Bar is visible, check Progress Menu
    mnuStatusBar.Checked = stbStatus.Visible        ' If Status Bar is visible, check Status Menu
    mnuToolbar.Checked = tlbToolbar.Visible         ' If Toolbar is visible, check Toolbar Menu
    Me.Caption = GetFormCaption(Me.Name) & " - " & strFileName
    Form_Resize
End Sub

Private Sub cmdAdd_Click()
' This subroutine adds the label text and quantity
' to the bottom of the list that the user had just entered
' and then it had TotalQty re-calculate the toal number of labels.
    
    '200508 - NF
    'Cannot allow blanks
    If txtLabelText.Text = "" Then
        txtLabelText.SetFocus
        MsgBox "Must enter label text before adding to the list.", vbOKOnly + vbInformation
        Exit Sub
    End If
    If Val(txtLabelQty.Text) > 0 Then
        lsvLabels.ListItems.Add , , txtLabelText.Text, 23, 23
        lsvLabels.ListItems(lsvLabels.ListItems.Count).Selected = True
        lsvLabels.SelectedItem.SubItems(1) = Val(txtLabelQty.Text)
        lsvLabels.SelectedItem.EnsureVisible
        lsvLabels.ListItems(lsvLabels.ListItems.Count).Selected = False
        txtLabelText.Text = ""
        txtLabelText.SetFocus
        lblUserMsg.Caption = ""
        blnSaved = False
    End If
    TotalQty
End Sub

Private Sub cmdAdd_GotFocus()
' This subroutine disables the cut, copy, paste, and delete
' buttons and menus.

    mnuCopy.Enabled = False
    tlbToolbar.Buttons(12).Enabled = False
    mnuCut.Enabled = False
    tlbToolbar.Buttons(11).Enabled = False
    mnuDelete.Enabled = False
    tlbToolbar.Buttons(10).Enabled = False
    mnuPaste.Enabled = False
    mnuPopupPaste.Enabled = False
    tlbToolbar.Buttons(13).Enabled = False
End Sub

Private Sub filMakeFieldTags_Click()
        Dim p As Printer

        frmLogin.Show vbModal
        If Not frmLogin.LoginSucceeded Then Exit Sub
        
        #If DebugMode = 1 Then
            For Each p In Printers
            
                If InStr(p.DeviceName, "8000n") > 0 Then
                    Set Printer = p
                    Exit For
            
                End If
            
            Next p
            '===
        
            frmField.Show vbModal, Me
        
        #Else
        
            Load frmField
            If SetFieldTagsForm(frmField.hwnd) Then
                frmField.Show vbModal, Me
                Call SetTerminalLabelsForm(frmMain.hwnd)
            Else
                MsgBox "SetFieldTagsForm Failed.  Cannot set form to Field Tag size.", vbOKOnly
            End If
        
        #End If
    
    
End Sub

Private Sub Form_Load()

    'Check if proper printer is installed
    If Not CheckProperPrinter Then End

'Set Conditional Compilation constant to use PrintPreview
#If DebugMode = 0 Then
    frmMain.mnuPrintPreview.Visible = False
#ElseIf DebugMode = 1 Then
    frmMain.mnuPrintPreview.Visible = True
#End If

'Field Tags can be printed from any program version,
'   as long as the login password is known (Vicki R. should know it)
filMakeFieldTags.Visible = True


' This subroutine is called on startup.
' It reads in all the settings from the registry.
' If the previous settings do not exist, default
' settings are used instead.
    

    On Error Resume Next
    
    Set ptrDefault = Printer
    blnSaved = True                         ' File is not saved
    blnAllowPaste = False                   ' Paste is not allowed
    blnPrintAll = True                      ' Print all labels by default
    intCopies = 1                           ' Set the number of copies to 1 by default
    
    ' Load previous settings from the registry
    frmMain.WindowState = GetSetting("KASA", "ShopLabels", "WindowState", 0)
    frmMain.Top = GetSetting("KASA", "ShopLabels", "Top", 1440)
    frmMain.Left = GetSetting("KASA", "ShopLabels", "Left", 2175)
    frmMain.Height = GetSetting("KASA", "ShopLabels", "Height", 6120)
    frmMain.Width = GetSetting("KASA", "ShopLabels", "Width", 8500)
    tlbToolbar.Visible = GetSetting("KASA", "ShopLabels", "Toolbar", True)
    stbStatus.Visible = GetSetting("KASA", "ShopLabels", "StatusBar", True)
    prgProgress.Visible = GetSetting("KASA", "ShopLabels", "ProgBar", True)
    intDefaultQty = GetSetting("KASA", "ShopLabels", "DefaultQty", 1)
    
    blnLockFormat = GetSetting("KASA", "ShopLabels", "LockFormat", False)
    
    '=======================
    'CHG 200510 N.F.
    '   Only storing the desired label format in registry,
    '   and loading actual label specs from the formats.dat file, which
    '   should have current version of label specs.
    Dim sTemp(10) As String
    If FeatureMode = WIRE_LABELS Then
        'default it locked
        blnLockFormat = True
        strLabelFormat = GetSetting("KASA", "ShopLabels", "WireLabelFormat", "Wire Labels - Optical - Autosize")
        If strLabelFormat = "Panduit PTR3 AutoSize" Then
            'legacy format mapped to new name
            strLabelFormat = "Wire Labels - Optical - Autosize"
        End If
        Debug.Print SelectNewFormat(strLabelFormat, sTemp)
    ElseIf FeatureMode = TERMINAL_STRIPS Then
        strLabelFormat = GetSetting("KASA", "ShopLabels", "TerminalLabelFormat", "1. CA1")
        Debug.Print SelectNewFormat(strLabelFormat, sTemp)
    ElseIf FeatureMode = STOCORD_LABELS Then
        'default it locked
        blnLockFormat = True
        strLabelFormat = GetSetting("KASA", "ShopLabels", "StoCordLabelFormat", "StoCord Labels - Optical - Autosize")
        Debug.Print SelectNewFormat(strLabelFormat, sTemp)
    Else
        'New Feature Mode?
        Debug.Assert False
        MsgBox "Program Feature Mode has not been implemented : " & FeatureMode & vbCr & vbCr & "Program is running as Wire Labels Program.", vbOKOnly
        'Default to Wire Labels
        FeatureMode = WIRE_LABELS
        strLabelFormat = GetSetting("KASA", "ShopLabels", "WireLabelFormat", "Wire Labels - Optical - Autosize")
    End If
    strLabelFormat = Trim(strLabelFormat)
    If SelectNewFormat(strLabelFormat, sTemp) Then
        '=======================
        ' Set the Label Specs to the new/selected label format
        '=======================
        ' All defined in globals
        sngTopMargin = FormatLabelDimension(sTemp(0))
        sngLeftMargin = FormatLabelDimension(sTemp(1))
        sngWidth = FormatLabelDimension(sTemp(2))
        sngHeight = FormatLabelDimension(sTemp(3))
        sngSpacingTB = FormatLabelDimension(sTemp(4))
        sngSpacingRL = FormatLabelDimension(sTemp(5))
        intLines = Int(Val(sTemp(6)))
        intLabelsPerRow = Int(Val(sTemp(7)))
        intOptical = Val(sTemp(8))
        intAutoSize = Val(sTemp(9))
    Else
        Debug.Assert False
    End If
'    sngTopMargin = GetSetting("KASA", "ShopLabels", "TopMargin", 0.5)
'    sngLeftMargin = GetSetting("KASA", "ShopLabels", "LeftMargin", 0.5)
'    sngWidth = GetSetting("KASA", "ShopLabels", "LabelWidth", 1)
'    sngHeight = GetSetting("KASA", "ShopLabels", "LabelHeight", 0.75)
'    sngSpacingTB = GetSetting("KASA", "ShopLabels", "SpacingTB", 1.5)
'    sngSpacingRL = GetSetting("KASA", "ShopLabels", "spacingRL", 1)
'    intLines = GetSetting("KASA", "ShopLabels", "Lines", 3)
'    intLabelsPerRow = GetSetting("KASA", "ShopLabels", "LabelsPerRow", 4)
'    intOptical = GetSetting("KASA", "ShopLabels", "Optical", 0)
'    intAutoSize = GetSetting("KASA", "ShopLabels", "AutoSize", 0)
    '=======================
    
    With cdlDialog
        .FontName = GetSetting("KASA", "ShopLabels", "FontName", "Arial")
        .FontSize = GetSetting("KASA", "ShopLabels", "FontSize", 8.25)
        .FontBold = GetSetting("KASA", "ShopLabels", "FontBold", True)
        .FontItalic = GetSetting("KASA", "ShopLabels", "FontItalic", False)
        .FontStrikethru = GetSetting("KASA", "ShopLabels", "Strikethrough", False)
        .FontUnderline = GetSetting("KASA", "ShopLabels", "Underline", False)
    End With
    With lsvLabels.Font                             ' Change Current Settings To New Font Selection
        .Name = cdlDialog.FontName
        .Bold = cdlDialog.FontBold
        .Italic = cdlDialog.FontItalic
        .Size = cdlDialog.FontSize
        .Strikethrough = cdlDialog.FontStrikethru
        .Underline = cdlDialog.FontUnderline
    End With
    
    
    ' Check to see if there are any arguments on the command line
    If Trim$(Command$) = "" Or Dir(Mid$(Trim$(Command$), 2, Len(Trim$(Command$)) - 2)) = "" Then
        ' No arguments, so just start with a blank list
        ' Set up variables as if we were starting a new label list
        mnuNew_Click
    Else
        ' There are arguments on the command line in the form of a
        ' File name, so load the file into our program instead of a blank list
        OpenCommandLine
    End If
    
    ' Turn on NumLock
 '   GetKeyboardState kbArray
 '   kbArray.kbByte(&H90) = 1
 '   SetKeyboardState kbArray
    
    ' Turn on CapsLock
 '   GetKeyboardState kbArray
 '   kbArray.kbByte(&H14) = 1
 '   SetKeyboardState kbArray
End Sub



Private Sub Form_Paint()
' This subroutine calls the subroutine to update the
' menus that have check mark possibilities on them.

    CheckVisualStatus                               ' Update checked menus
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' This subroutine is called when the user tries to exit
' out of the program.  It checks to see if the file has
' been saved.  If it has, it just lets the program exit.
' If the file has not been saved, it allows the user to
' save the file before exiting.  If the user chooses cancel,
' it stops the program from exiting.

On Error Resume Next
    
    Dim Result As Long
    
    ' Check to see if label list has been saved or not
    If blnSaved = False Then
        ' File has not been saved, prompt user to save before exiting
        Result = MsgBox("Label list has changed.  Save changes?", vbYesNoCancel + vbExclamation, "Save changes?")
        If Result = vbYes Then
            ' They want to save the file, start save routine
            mnuSave_Click
        ElseIf Result = vbCancel Then
            ' They wanted to save, but ended up clicking Cancel
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Resize()
' This subroutine makes sure that all the controls on the
' window are within the limits defined in this subroutine.
' If anything is outside the limits, it adjusts it accordingly.

    Dim intDiff As Long
    
    If frmMain.WindowState = vbMinimized Then       ' If minimized, don't resize objects on form
        Exit Sub
    End If
    
    If frmMain.Height < 3100 Then
        ' They're trying to resize it smaller than allowed, so
        ' automatically release their mouse button to prevent it
        PostMessage frmMain.hwnd, WM_LBUTTONUP, 0&, 0&
        frmMain.Height = 3100
    End If
    
    
    If frmMain.Width < 8200 Then
        ' They're trying to resize it smaller than allowed, so
        ' automatically release their mouse button to prevent it
        PostMessage frmMain.hwnd, WM_LBUTTONUP, 0&, 0&
        frmMain.Width = 8200
    End If
    
    ' The difference between the form height and the list height is already 1755
    intDiff = 1775
    
    ' If the status bar is visible, add on the height to the height difference
    If stbStatus.Visible Then
        intDiff = intDiff + stbStatus.Height        ' Adjust ListView Height to accomodate empty window space
    End If
    
    ' If the progress bar is visible, add on the height to the height difference
    If prgProgress.Visible Then
        intDiff = intDiff + prgProgress.Height      ' Adjust ListView Height to accomodate empty window space
    End If
    
    lsvLabels.Width = frmMain.Width - 100           ' Resize ListView Width
    lsvLabels.Height = frmMain.Height - intDiff     ' Resize Listview Height
    lsvLabels.ColumnHeaders(1).Width = 0.8 * lsvLabels.Width    ' Resize column widths
    lsvLabels.ColumnHeaders(2).Width = 0.15 * lsvLabels.Width   ' Resize column widths
    lblUserMsg.Width = frmMain.Width - 3720         ' Resize Label Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
' This subroutine is called just after the Form_QueryUnload
' subroutine if the user does not choose cancel, and it executes
' just before completely exiting.  It saves all the current
' settings into the registry so that they can be loaded up
' on startup the next time the program is run.

    On Error Resume Next
    
    ' Clear the clipboard so that it isn't using any memory
    Clipboard.Clear
    
    Set Printer = ptrDefault
    
    ' Unload all forms
    Unload frmAbout
    Unload frmAlpha
    Unload frmBrowser
    Unload frmDecimal
    Unload frmFormat
    Unload frmInsertChar
    Unload frmPLC
    Unload frmPrint
    Unload frmPrinterSetup
    Unload frmSLC
    
    ' Save current settings to the registry
    If frmMain.WindowState = vbMinimized Then
        ' Don't let the program start up minimized, so if it is
        ' minimized when the user exits, set the startup to normal
        SaveSetting "KASA", "ShopLabels", "WindowState", vbNormal
    Else
        SaveSetting "KASA", "ShopLabels", "WindowState", frmMain.WindowState
    End If
    
    If frmMain.WindowState = vbNormal Then
        SaveSetting "KASA", "ShopLabels", "Top", frmMain.Top
        SaveSetting "KASA", "ShopLabels", "Left", frmMain.Left
        If frmMain.Height < 3200 Then
            SaveSetting "KASA", "ShopLabels", "Height", 3200
        Else
            SaveSetting "KASA", "ShopLabels", "Height", frmMain.Height
        End If
        If frmMain.Width < 8300 Then
            SaveSetting "KASA", "ShopLabels", "Width", 8300
        Else
            SaveSetting "KASA", "ShopLabels", "Width", frmMain.Width
        End If
    End If
    
    With cdlDialog
        SaveSetting "KASA", "ShopLabels", "FontName", .FontName
        SaveSetting "KASA", "ShopLabels", "FontSize", .FontSize
        SaveSetting "KASA", "ShopLabels", "FontBold", .FontBold
        SaveSetting "KASA", "ShopLabels", "FontItalic", .FontItalic
        SaveSetting "KASA", "ShopLabels", "Strikethrough", .FontStrikethru
        SaveSetting "KASA", "ShopLabels", "Underline", .FontUnderline
    End With
    
    SaveSetting "KASA", "ShopLabels", "Toolbar", tlbToolbar.Visible
    SaveSetting "KASA", "ShopLabels", "StatusBar", stbStatus.Visible
    SaveSetting "KASA", "ShopLabels", "ProgBar", prgProgress.Visible
    SaveSetting "KASA", "ShopLabels", "DefaultQty", intDefaultQty
    SaveSetting "KASA", "ShopLabels", "LockFormat", blnLockFormat
    If FeatureMode = WIRE_LABELS Then
        SaveSetting "KASA", "ShopLabels", "WireLabelFormat", strLabelFormat
    ElseIf FeatureMode = TERMINAL_STRIPS Then
        SaveSetting "KASA", "ShopLabels", "TerminalLabelFormat", strLabelFormat
    ElseIf FeatureMode = STOCORD_LABELS Then
        SaveSetting "KASA", "ShopLabels", "StoCordLabelFormat", strLabelFormat
    Else
        'New Feature Mode?
        Debug.Assert False
    End If
'    SaveSetting "KASA", "ShopLabels", "TopMargin", sngTopMargin
'    SaveSetting "KASA", "ShopLabels", "LeftMargin", sngLeftMargin
'    SaveSetting "KASA", "ShopLabels", "LabelWidth", sngWidth
'    SaveSetting "KASA", "ShopLabels", "LabelHeight", sngHeight
'    SaveSetting "KASA", "ShopLabels", "SpacingTB", sngSpacingTB
'    SaveSetting "KASA", "ShopLabels", "SpacingRL", sngSpacingRL
'    SaveSetting "KASA", "ShopLabels", "Lines", intLines
'    SaveSetting "KASA", "ShopLabels", "LabelsPerRow", intLabelsPerRow
'    SaveSetting "KASA", "ShopLabels", "Optical", intOptical
'    SaveSetting "KASA", "ShopLabels", "AutoSize", intAutoSize

End Sub

Private Sub lsvLabels_AfterLabelEdit(Cancel As Integer, NewString As String)
' This subroutine executes after the user has modified the
' text of one of the labels.  It first checks to see if the
' user actually changed the text, and if they did, it updates
' the status bar with the new text, and strips any spaces and
' null characters off the text before saving it.

    If Trim$(NewString) <> strOldLabel Then
        blnSaved = False
        NewString = Trim$(NewString)
        stbStatus.Panels(2).Text = "Label Text: " & NewString
    End If
End Sub

Private Sub lsvLabels_BeforeLabelEdit(Cancel As Integer)
' This subroutine executes just before it allows the user to
' edit the text of the label.  It stores the current text
' into a variable and checks to see if it changed in the
' lsvLabels_AfterLabelEdit subroutine.

    strOldLabel = Trim$(lsvLabels.SelectedItem.Text)
    TotalQty
End Sub

Private Sub lsvLabels_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' This subroutine sorts the column that was clicked in the
' opposite order that it was last sorted by.  If the list
' was never sorted, it sorts it in ascending order.

    On Error Resume Next
    
    blnSaved = False                                ' Allow user to save changes
    
    lsvLabels.Sorted = True                         ' Sort items
    lsvLabels.SortKey = ColumnHeader.Index - 1      ' Sort items by column number
    
    If lsvLabels.SortOrder = lvwAscending Then      ' Change sort order
        lsvLabels.SortOrder = lvwDescending
    Else
        lsvLabels.SortOrder = lvwAscending
    End If
    
    lsvLabels.Sorted = False                        ' Turn off Sort so new items stay at bottom
End Sub

Private Sub lsvLabels_DblClick()

'Doesn't work!
'    'This permits them to change the label
'    lsvLabels.StartLabelEdit
    
End Sub

Private Sub lsvLabels_GotFocus()
' This subroutine checks to see if anything is selected.
' If there is, enable cut, copy, paste, and delete otherwise
' disable them.

    On Error Resume Next
    If lsvLabels.SelectedItem Is Nothing Then
        ' No item selected, disable cut, copy, paste, delete, etc.
        mnuCopy.Enabled = False
        mnuCut.Enabled = False
        mnuPaste.Enabled = False
        mnuDelete.Enabled = False
        mnuPopupCopy.Enabled = False
        mnuPopupCut.Enabled = False
        mnuPopupPaste.Enabled = False
        mnuPopupDelete.Enabled = False
        tlbToolbar.Buttons(10).Enabled = False
        tlbToolbar.Buttons(11).Enabled = False
        tlbToolbar.Buttons(12).Enabled = False
        tlbToolbar.Buttons(13).Enabled = False
    Else
        ' Item selected, enable cut, copy, paste, delete, etc.
        mnuCopy.Enabled = True
        mnuCut.Enabled = True
        mnuDelete.Enabled = True
        mnuPopupCopy.Enabled = True
        mnuPopupCut.Enabled = True
        mnuPopupDelete.Enabled = True
        tlbToolbar.Buttons(10).Enabled = True
        tlbToolbar.Buttons(11).Enabled = True
        tlbToolbar.Buttons(12).Enabled = True
        ' Check to see if we did copy something to the clipboard
        ' before enabling the paste button
        If blnAllowPaste = True Then
            mnuPaste.Enabled = True
            mnuPopupPaste.Enabled = True
            tlbToolbar.Buttons(13).Enabled = True
        End If
    End If
End Sub

Private Sub lsvLabels_ItemClick(ByVal Item As MSComctlLib.ListItem)
' This subroutine executes whenever an item is clicked in the
' list view.  It first checks to see if they actually clicked
' on a valid entry, then allows the user to cut, copy, delete,
' and depending if we've copied anything, paste also.  It also
' updates the status bar to show the information concerning the
' current selected item.

    On Error GoTo ErrorHandler
    
    ' Enable cut, copy, and delete.
    ' Enable paste if we have copied a label into the clipboard.
    mnuCopy.Enabled = True
    tlbToolbar.Buttons(12).Enabled = True
    mnuCut.Enabled = True
    tlbToolbar.Buttons(11).Enabled = True
    If blnAllowPaste = True Then
        mnuPaste.Enabled = True
        mnuPopupPaste.Enabled = True
        tlbToolbar.Buttons(13).Enabled = True
    End If
    mnuDelete.Enabled = True
    tlbToolbar.Buttons(10).Enabled = True
    
    ' Change each panel in the status bar
    stbStatus.Panels.Item(1).Text = "Label: " & lsvLabels.SelectedItem.Index & "/" & lsvLabels.ListItems.Count
    If Trim$(lsvLabels.SelectedItem.Text) = "" Then
        stbStatus.Panels.Item(2).Text = "(Blank Label)"
    Else
        stbStatus.Panels.Item(2).Text = "Label Text: " & lsvLabels.SelectedItem.Text
    End If
    stbStatus.Panels.Item(3).Text = "Qty: " & Trim$(lsvLabels.SelectedItem.SubItems(1))
    
    Exit Sub
ErrorHandler:
    ' Change each panel in the status bar
    stbStatus.Panels.Item(1).Text = "Label: "
    stbStatus.Panels.Item(2).Text = "Label Text: "
    stbStatus.Panels.Item(3).Text = "Qty: "
End Sub

Private Sub lsvLabels_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' This subroutine checks to see if the user is right clicking
' on one of the items in the list view.  If they are, it just
' displays the popup menu at the current mouse location.

    On Error Resume Next
    
    ' If they right clicked on a label, display popup menu
    If Button = 2 Then
        Set lsvLabels.DropHighlight = lsvLabels.HitTest(X, Y)
        If lsvLabels.DropHighlight Is Nothing Then
            Set lsvLabels.DropHighlight = Nothing
            Exit Sub
        Else
            PopupMenu mnuPopup, vbPopupMenuLeftButton
        End If
    End If
    Set lsvLabels.DropHighlight = Nothing
End Sub

Private Sub mnuAbout_Click()
' This subroutine shows the about dialog box
    frmAbout.Show 1
End Sub

Private Sub mnuAlpha_Click()
' This subroutine turns on the CapsLock and shows the
' alphabetic sequence dialog box
    
    ' Make sure the CapsLock is On
    'GetKeyboardState kbArray
    'kbArray.kbByte(&H14) = 1
    'SetKeyboardState kbArray
    
    ' Show the alphabetic sequence dialog
    frmAlpha.Show 1
End Sub

Private Sub mnuChgAllQty_Click()
' This subroutine selects all the labels in the list
' and changes their quantities to the number chosen
' in the quantity dialog box.

    ' If there are no labels in the list, just exit this subroutine
    If lsvLabels.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    ' Select all the labels
    mnuSelectAll_Click
    
    ' Run the Change Quantity subroutine
    mnuChgQty_Click
End Sub

Private Sub mnuChgQty_Click()
' This subroutine allows the user to change the quantity
' of the currently selected label.
    
    Dim intCurQty As Long
    Dim strNewQty As String
    Dim itmItem As ListItem
    
    ' Check to see if anything is selected.
    ' If there isn't, just exit this subroutine
    If lsvLabels.SelectedItem Is Nothing Then
        Exit Sub
    End If
        
    intCurQty = lsvLabels.SelectedItem.SubItems(1)  ' Set Current Quantity
    
    ' Display InputBox with default as Current Quantity
    strNewQty = InputBox("Enter new quantity for label:", "Change Quantity", intCurQty)
    
    ' If they typed in a valid number and they didn't hit Cancel, change to new Quantity
    If Val(strNewQty) > 0 Then
        For Each itmItem In lsvLabels.ListItems
            If itmItem.Selected = True Then
                itmItem.SubItems(1) = Trim$(Val(strNewQty))
            End If
        Next itmItem
        blnSaved = False
        TotalQty
    End If
    lsvLabels_ItemClick lsvLabels.SelectedItem
End Sub

Private Sub mnuContents_Click()
' This subroutine shows the Help File in it's own window.

    frmBrowser.Show
End Sub

Private Sub mnuCopy_Click()
' This subroutine copies the currently selected labels and their
' respective quantities into the clipboard using the predetermined
' format that is shown in this subroutine.  It also allows the
' paste menus and the paste toolbutton to be clicked if the user
' has a label highlighted in the list view.
    
    Dim strSelection As String
    Dim intLoop As Long
    
    ' Set String to Initialization character which also
    ' helps check when we paste to make sure the data is
    ' originally from our program
    strSelection = Chr(255) & App.Title & Chr(255)
    
    ' Add each selected label's text and an initialization
    ' character to the end of the selection string to seperate'
    ' the different label texts
    For intLoop = 1 To lsvLabels.ListItems.Count
        If lsvLabels.ListItems(intLoop).Selected = True Then
            strSelection = strSelection & lsvLabels.ListItems(intLoop).Text & Chr(255)
        End If
    Next intLoop
    
    ' Add the tilde character to seperate the label text
    ' from the label quantities
    strSelection = strSelection & "~"
    
    ' Add each selected label's quantity and an initialization
    ' character to the end of the selection string to seperate
    ' the different label quantities
    For intLoop = 1 To lsvLabels.ListItems.Count
        If lsvLabels.ListItems(intLoop).Selected = True Then
            strSelection = strSelection & lsvLabels.ListItems.Item(intLoop).SubItems(1) & Chr(255)
        End If
    Next intLoop
    
    ' Clear current clipboard contents and set the clipboard
    ' text contents to the selection string we just made
    Clipboard.Clear
    Clipboard.SetText strSelection
    
    ' Set boolean variable for pasting to true to allow
    ' the user to paste now and enable all the paste menus
    ' and buttons on our form
    blnAllowPaste = True
    mnuPaste.Enabled = True
    mnuPopupPaste.Enabled = True
    tlbToolbar.Buttons(13).Enabled = True
End Sub

Private Sub mnuCut_Click()
' This subroutine calls the mnuCopy_Click and the
' mnuDelete_Click procedures, because this subroutine
' would be duplicating those exactly.  It copies the
' labels to the clipboard, and then deletes them.
    
    mnuCopy_Click                           ' Copy label text and qty
    mnuDelete_Click                         ' Delete label entry
End Sub

Private Sub mnuDecimal_Click()
' This subroutine shows the decimal sequence dialog box.

    frmDecimal.Show 1
End Sub

Private Sub mnuDefaultQty_Click()
' This subroutine allows the user to change the default quantity
' of labels that is used when the program starts, or when a file
' is imported in from a TXT file.  It also changes the current
' number in the label quantity text box to the new quantity.

    Dim strTemp As String
    
    ' Prompt user for number of default labels
    strTemp = InputBox("Enter default number of labels:", "Set Default Quantity", intDefaultQty)
    
    ' If they canceled, or typed in a value less than 1
    ' then just exit sub without changing default value
    If Val(strTemp) < 1 Then
        Exit Sub
    End If
    
    ' Set default quantity to the new value
    intDefaultQty = Val(strTemp)
    
    ' Change the text for the next new label to the new
    ' default quantity that was just selected
    txtLabelQty.Text = intDefaultQty
End Sub

Private Sub mnuDelete_Click()
' This subroutine goes through the list backwards, checking
' for the selected records that need to be deleted.  It
' removes them one-by-one from the list as it finds them.
' Then it updates the total quantity of labels displayed
' in the status bar.

    Dim intLoop As Long
    
    If lsvLabels.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    ' Delete the items in reverse order, so that we don't
    ' accidently delete a wrong item, because when you delete
    ' an item, the indexes for the following items change.
    ' Deleting them in reverse order keeps us from messing up.
    For intLoop = lsvLabels.ListItems.Count To 1 Step -1
        If lsvLabels.ListItems(intLoop).Selected = True Then
            lsvLabels.ListItems.Remove intLoop
        End If
    Next intLoop
    
    ' Change each panel in the status bar
    stbStatus.Panels.Item(1).Text = "Label: 0/" & lsvLabels.ListItems.Count
    stbStatus.Panels.Item(2).Text = "(No Label Selected)"
    stbStatus.Panels.Item(3).Text = "Qty: 0"

    TotalQty                                ' Re-Total labels
End Sub

Private Sub mnuDeleteChar_Click()
' This subroutine deletes a character position from the
' currently selected labels in the list.

    On Error Resume Next
    
    Dim itmItem As ListItem
    Dim strTemp As String
    Dim intTemp As Long
    
    ' If there is nothing selected, just exit this subroutine.
    If lsvLabels.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    ' Input the character position to delete from the selected labels
    strTemp = InputBox("Enter character position from left to delete:", "Delete Character", 1)
    
    ' If they didn't type anything or clicked Cancel, just exit subroutine
    If Trim$(strTemp) = "" Then
        Exit Sub
    End If
    
    ' Convert the string value to an integer
    intTemp = Int(Val(strTemp))
    
    ' If they selected any position less than one, it is an
    ' invalid position, let the user know, and open the dialog
    ' box for input again.
    If intTemp < 1 Then
        MsgBox "Invalid character position!", vbOKOnly + vbInformation, "Delete Character"
        mnuDeleteChar_Click
        Exit Sub
    End If
    
    ' Go through list deleting the selected character position from each label
    For Each itmItem In lsvLabels.ListItems
        If itmItem.Selected = True Then
            strTemp = Left$(itmItem.Text, intTemp - 1)
            strTemp = strTemp & Right$(itmItem.Text, Len(itmItem.Text) - intTemp)
            itmItem.Text = strTemp
        End If
    Next itmItem
    
    ' List has changed, and hasn't been saved yet.
    blnSaved = False
    lsvLabels_ItemClick lsvLabels.SelectedItem
End Sub

Private Sub mnuExit_Click()
' This subroutine tells the program to begin the procedures
' that are executed just before the program exits, and then
' it exits the program.

    Unload Me                                       ' Exit Program
End Sub

Private Sub mnuExport_Click()
' This subroutine allows the user to export the label
' list to a file that can be read in any text viewer.
' It will not export quantities, only label texts.

    Dim intResults As Integer
    Dim lstListItem As ListItem

    On Error Resume Next
    
    ' If we don't have any labels in the list, just exit sub
    If lsvLabels.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    cdlDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    cdlDialog.Flags = 0
    cdlDialog.ShowSave
    cdlDialog.Flags = cdlOFNOverwritePrompt
    
    If Err.Number <> cdlCancel And Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & " - " & Err.Description
        Exit Sub
    ElseIf Err.Number = cdlCancel Then
        Exit Sub
    End If
    
    If Len(Dir(cdlDialog.FileName)) <> 0 Then
        intResults = MsgBox("File already exists!" & vbCrLf & vbCrLf & "Yes - Overwrite File" & vbCrLf & "No - Append to File" & vbCrLf & "Cancel - Select Another Filename", vbYesNoCancel, "Overwrite File?")
        If intResults = vbYes Then
            Open cdlDialog.FileName For Output As #1
        ElseIf intResults = vbNo Then
            Open cdlDialog.FileName For Append As #1
        Else
            mnuExport_Click
            Exit Sub
        End If
    Else
        Open cdlDialog.FileName For Output As #1
    End If
    
    prgProgress.Value = 0
    prgProgress.Max = lsvLabels.ListItems.Count
    
    For Each lstListItem In lsvLabels.ListItems
        Print #1, lstListItem.Text
        prgProgress.Value = lstListItem.Index + 1
    Next 'lstListItem
    
    prgProgress.Value = 0
    
    Close #1
End Sub

Private Sub mnuFind_Click()
' This subroutine starts searching the label list for the
' first item in the list that matches what the user inputs
' into the input box that is displayed

    On Error Resume Next
    
    Dim itmFind As ListItem
    Dim intCnt As Long
    
    ' If we don't have any labels in the list, just exit sub
    If lsvLabels.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    ' Prompt user for search string text
    ' If there is a label selected, set the default search
    ' string text to the text of that label, otherwise set
    ' the search string text to nothing
    If lsvLabels.SelectedItem Is Nothing Then
        strFind = InputBox("Label text to search for:", "Find")
    Else
        strFind = InputBox("Label text to search for:", "Find", lsvLabels.SelectedItem.Text)
    End If
    
    ' Check to see if the user clicked cancel or didn't type
    ' anything into the input box.  If they didn't just exit.
    If Trim$(strFind) = "" Then
        Exit Sub
    End If
    
    ' Have the list view find the next item for us automatically
    Set itmFind = lsvLabels.FindItem(strFind)
    
    ' If nothing was found, let the user know and exit the sub.
    If itmFind Is Nothing Then
        MsgBox Chr(34) & strFind & Chr(34) & " Not Found!", vbInformation, "Find"
        Exit Sub
    End If
    
    ' Unselect all the items, so that the currently selected
    ' Item still isn't selected when we select the found item.
    For intCnt = 1 To lsvLabels.ListItems.Count
        lsvLabels.ListItems.Item(intCnt).Selected = False
    Next intCnt
    
    ' Select the found item, make sure it's visible, and update
    ' the status bar for the item we just found.
    lsvLabels.SelectedItem = itmFind
    lsvLabels.SelectedItem.EnsureVisible
    lsvLabels_ItemClick lsvLabels.SelectedItem
    lsvLabels.SetFocus
End Sub

Private Sub mnuFindNext_Click()
' This subroutine starts searching the label list for the
' next item in the list that matches what the user inputed
' into the input box that was displayed when they selected
' find the first time

    On Error Resume Next
    
    Dim itmFind As ListItem
    Dim intCnt As Long
    
    ' If we haven't performed a search yet, show the find
    ' dialog box instead and exit this sub.
    If strFind = "" Then
        mnuFind_Click
        Exit Sub
    End If
    
    ' Let the list view find the next item from the currently
    ' selected item.
    Set itmFind = lsvLabels.FindItem(strFind, , lsvLabels.SelectedItem.Index + 1)
    
    ' If nothing was found, then notify the user and exit this sub.
    If itmFind Is Nothing Then
        MsgBox Chr(34) & strFind & Chr(34) & " Not Found!", vbInformation, "Find"
        Exit Sub
    End If
    
    ' Unselect all the items, so that the currently selected
    ' Item still isn't selected when we select the found item.
    For intCnt = 1 To lsvLabels.ListItems.Count
        lsvLabels.ListItems.Item(intCnt).Selected = False
    Next intCnt
    
    ' Select the found item, make sure it's visible, and update
    ' the status bar for the item we just found.
    lsvLabels.SelectedItem = itmFind
    lsvLabels.SelectedItem.EnsureVisible
    lsvLabels_ItemClick lsvLabels.SelectedItem
    lsvLabels.SetFocus
End Sub

Private Sub mnuFont_Click()
' This subroutine reads the current font settings into the
' font dialog box, and then displays the dialog box.  When
' the user has selected the new font, it updates the list
' view to the new font settings.

    On Error GoTo ErrorHandler
    
    With cdlDialog                                  ' Set Up Font Selection To Current Settings
        .FontName = lsvLabels.Font.Name
        .FontBold = lsvLabels.Font.Bold
        .FontItalic = lsvLabels.Font.Italic
        .FontSize = lsvLabels.Font.Size
        .FontStrikethru = lsvLabels.Font.Strikethrough
        .FontUnderline = lsvLabels.Font.Underline
        .Flags = cdlCFPrinterFonts
    End With
    
    cdlDialog.ShowFont                              ' Show Font Selection Dialog Box
    
    With lsvLabels.Font                             ' Change Current Settings To New Font Selection
        .Name = cdlDialog.FontName
        .Bold = cdlDialog.FontBold
        .Italic = cdlDialog.FontItalic
        .Size = cdlDialog.FontSize
        .Strikethrough = cdlDialog.FontStrikethru
        .Underline = cdlDialog.FontUnderline
    End With
    
    Exit Sub                                        ' Bypass Error Handler
    
ErrorHandler:
    If Err.Number = 32755 Or Err.Number = 24574 Then ' If Cancel is Selected
        Exit Sub
    Else                                            ' Any other errors
        MsgBox "Error: " & Err.Number & " - " & Err.Description
    End If
End Sub

Private Sub mnuImport_Click()
' This subroutine allows the user to import a file
' exported by AutoCAD as a TXT file.  As it reads in
' the file, it sets the quantity of labels to the
' default quantity set by the user.

    On Error GoTo ErrorHandler
    
    Dim intCount As Long
    Dim Result As Long
    Dim strTemp As String
    
    ' Set the dialog box filter to *.TXT and  *.* Files
    ' Then show the dialog box for opening files.
    cdlDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    cdlDialog.Flags = cdlOFNHideReadOnly
    cdlDialog.ShowOpen                              ' Show Open File Dialog Box
    
    'If they didn't select a file just exit sub.
    If cdlDialog.FileName = "" Then
        Exit Sub
    End If
    
    ' Show an hourglass mouse pointer
    frmMain.MousePointer = 11
    
    ' Open the file for input so we can read the values
    Open cdlDialog.FileName For Input As #1
    
    ' Read in each label until we reach the end of the file.
    While Not EOF(1)
        Line Input #1, strTemp
        intCount = intCount + 1
        lsvLabels.ListItems.Add , , Trim$(strTemp), 23, 23
        lsvLabels.ListItems(lsvLabels.ListItems.Count).SubItems(1) = intDefaultQty
    Wend
    
    ' Close the file
    Close #1
    
    ' Unselect the first item, which was selected by default
    ' by the list view
    UnSelect
    
    ' Show the normal mouse pointer again
    frmMain.MousePointer = 0
    
    ' Set focus to the text box so the user will be able
    ' to add a new label if they want to
    txtLabelText.SetFocus
    
    ' The label list has changed, so it has not been saved yet
    ' and re-total the number of labels
    blnSaved = False
    lsvLabels.Refresh
    lsvLabels.ListItems(lsvLabels.ListItems.Count).EnsureVisible
    TotalQty
    Exit Sub                                        ' Bypass Error Handler

ErrorHandler:
    If Err.Number = 32755 Then                      ' If Cancel is Selected
        Exit Sub
    ElseIf Err.Number = 75 Or Err.Number = 57 Then  ' Bad File/Disk - Error reading file
        MsgBox "The file is corrupt, possibly a bad disk!"
    Else                                            ' Any other errors
        MsgBox "Error: " & Err.Number & " - " & Err.Description
    End If
End Sub

Private Sub mnuInsertBlank_Click()
' This subroutine will insert 1 blank label

    On Error Resume Next
    
    ' If there are no labels, just add a blank label to the list
    ' otherwise, put the blank label where we selected, and shift
    ' the other labels down.
    If lsvLabels.ListItems.Count = 0 Then
        lsvLabels.ListItems.Add , , , 23, 23
        lsvLabels.ListItems.Item(lsvLabels.ListItems.Count).SubItems(1) = 1
    Else
        lsvLabels.ListItems.Add lsvLabels.SelectedItem.Index, , , 23, 23
        lsvLabels.ListItems.Item(lsvLabels.SelectedItem.Index - 1).SubItems(1) = 1
        lsvLabels.ListItems.Item(lsvLabels.SelectedItem.Index - 1).Selected = True
        lsvLabels.ListItems.Item(lsvLabels.SelectedItem.Index + 1).Selected = False
    End If
    
    lsvLabels_ItemClick lsvLabels.SelectedItem  ' Select the blank label.
    blnSaved = False                            ' List has changed
End Sub

Private Sub mnuInsertChar_Click()
' This subroutine shows the insert character dialog box.

    ' If there is no label selected, just exit this subroutine
    If lsvLabels.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    ' Show the insert character dialog box
    frmInsertChar.Show 1
    
    lsvLabels_ItemClick lsvLabels.SelectedItem
End Sub

Private Sub mnuLabelFormat_Click()
' This subroutine shows the label format dialog box.

    frmFormat.Show 1
End Sub

Private Sub mnuNew_Click()
' This subroutine confirms that the user wants to start a new
' label list and clear out the old one.  It also checks to see
' if the old list has been saved.  If it hasn't, it gives the
' user a chance to save it before clearing the old list out.

    On Error Resume Next
    
    Dim Result As Long

    ' Check the current filename, if there isn't one, then
    ' we just started the application and don't prompt the
    ' user if they are sure if they want to clear the list
    If strFileName <> "" Then
        Result = MsgBox("Are you sure you want to clear the label list?", vbYesNo + vbExclamation, "Clear List?")
            If Result = vbNo Then
            Exit Sub
        End If
    Else
        blnSaved = True                     ' Just started app - blank list - nothing to save
    End If
    
    ' If the list they currently have hasn't been saved,
    ' give the user a chance to save what they had.
    If blnSaved = False Then
        Result = MsgBox("Label list has changed.  Save changes?", vbYesNoCancel + vbExclamation, "Save changes?")
        If Result = vbYes Then
            mnuSave_Click
        ElseIf Result = vbCancel Then
            Exit Sub
        End If
    End If
    
    '200510 N.F.
    'If running in Wire Labels or StoCord Labels mode,
    '   no need to ask this question... format should already be loaded.
    If FeatureMode = WIRE_LABELS Or FeatureMode = STOCORD_LABELS Then
        If FeatureMode = WIRE_LABELS Then
        'Wire Labels - Optical - Autosize
            Debug.Assert InStr(1, strLabelFormat, "wire label", vbTextCompare) > 0
        ElseIf FeatureMode = STOCORD_LABELS Then
        'StoCord Labels - Optical - Autosize
            Debug.Assert InStr(1, strLabelFormat, "stocord label", vbTextCompare) > 0
        End If
    Else
        If Not blnLockFormat Then
            ' Check to see if they want to change the current label format.
            Result = MsgBox("The current label format is: " & Chr(34) & Trim$(strLabelFormat) & Chr(34) & "." & vbCrLf & "Do you want to change this?", vbYesNo + vbExclamation, "Change Format?")
            If Result = vbYes Then
                frmFormat.Show 1
            End If
        End If
    End If

    ' New list, so there is nothing to save, set the filename
    ' to just "Untitled", clear the list and set the quantity
    ' text box to the default quantity, clear the search string
    ' text, and re-total the quantity of labels for the status bar
    blnSaved = True
    strFileName = "Untitled"
    lsvLabels.ListItems.Clear
    txtLabelQty.Text = intDefaultQty
    lblUserMsg.Caption = ""
    strFind = ""
    mnuCopy.Enabled = False
    tlbToolbar.Buttons(12).Enabled = False
    mnuCut.Enabled = False
    tlbToolbar.Buttons(11).Enabled = False
    mnuDelete.Enabled = False
    tlbToolbar.Buttons(10).Enabled = False
    TotalQty
    stbStatus.Panels(1).Text = "Label: 0/0"
    stbStatus.Panels(2).Text = "(No Label Selected)"
    stbStatus.Panels(3).Text = "Qty: 0"
End Sub

Private Sub mnuOpen_Click()
' This subroutine allows the user to open a label list file
' that they have previously saved.  It also checks to see
' if they have saved the current list yet.  If not, it gives
' them a chance to save it before they open the file they
' just selected to open.

    On Error GoTo ErrorHandler
    
    Dim intCount As Long
    Dim Result As Long
    Dim strTemp1 As String
    Dim strTemp2 As String
    
    ' Allow the user to save the current label list before opening
    ' a label list that was saved to disk
    If blnSaved = False Then
        Result = MsgBox("Label list has changed.  Save changes?", vbYesNoCancel + vbExclamation, "Save changes?")
        If Result = vbYes Then
            mnuSave_Click
        ElseIf Result = vbCancel Then
            Exit Sub
        End If
    End If
    
    ' Set the dialog box filter to *.LBL and *.* files, then
    ' set the dialog box filename to the default name, and
    ' open the dialog box
    cdlDialog.Filter = "Label List Files (*.lbl)|*.lbl|All Files (*.*)|*.*"
    cdlDialog.Flags = cdlOFNHideReadOnly
    cdlDialog.FileName = strFileName
    cdlDialog.ShowOpen                              ' Show Open File Dialog Box
    
    ' If they didn't select a file, then just exit sub
    If cdlDialog.FileName = "" Then
        Exit Sub
    End If
    
    ' Add the file to the documents menu on the start menu
    SHAddToRecentDocs SHARD_PATH, cdlDialog.FileName
    
    ' Set the default filename to the file they just selected
    strFileName = cdlDialog.FileName
    
    ' Set the mouse pointer to an hourglass
    frmMain.MousePointer = 11
    
    ' Clear the label list
    lsvLabels.ListItems.Clear
    
    ' Open the file for input
    Open cdlDialog.FileName For Input As #1
    
    ' Add each item to the list from the data in the file
    ' and keep reading them in until we get to the end of
    ' the file
    While Not EOF(1)
        intCount = intCount + 1
        Line Input #1, strTemp1
        Line Input #1, strTemp2
        If Trim$(strTemp2) <> "" Then
            lsvLabels.ListItems.Add , , Trim$(strTemp1), 23, 23
            lsvLabels.ListItems(intCount).SubItems(1) = Trim$(strTemp2)
        Else
            intCount = intCount - 1
        End If
    Wend
    
    ' Close the file
    Close #1
        
    If Not blnLockFormat Then
        ' Check to see if they want to change the current label format.
        Result = MsgBox("The current label format is: " & Chr(34) & Trim$(strLabelFormat) & Chr(34) & "." & vbCrLf & "Do you want to change this?", vbYesNo + vbExclamation, "Change Format?")
        If Result = vbYes Then
            frmFormat.Show 1
        End If
    End If
    
    ' Unselect the first item that is selected by default
    lsvLabels.ListItems(1).Selected = False
    
    lblUserMsg.Caption = ""
    
    ' Set the mouse pointer back to the default
    frmMain.MousePointer = 0
    
    ' Give focus to the label text text box so the user
    ' can enter in any new labels if he/she wants to
    txtLabelText.SetFocus
    
    ' Clear the search string text
    strFind = ""

    ' File is the same as when it was opened, so no need to save
    blnSaved = True
    
    ' Re-Total the label quantites and update the status bar
    TotalQty
    
    Exit Sub                                        ' Bypass Error Handler

ErrorHandler:
    If Err.Number = 32755 Then                      ' If Cancel is Selected
        Exit Sub
    ElseIf Err.Number = 75 Or Err.Number = 57 Then  ' Bad File/Disk - Error reading file
        MsgBox "The file is corrupt, possibly a bad disk!"
    Else                                            ' Any other errors
        MsgBox "Error: " & Err.Number & " - " & Err.Description
    End If
End Sub

Private Sub mnuPaste_Click()
' This subroutine allows the user to paste the labels
' that they copied to the clipboard using the copy or
' cut commands.  It also checks to make sure that the
' labels being imported are of the correct format, so
' that the program cannot be crashed by invalid data.

    On Error Resume Next
    
    Dim strSelect As String
    Dim strTemp As String
    Dim intCount As Long
    Dim intCount2 As Long
    Dim intCnt As Long
    Dim intPasteStart As Long
        
    ' Read in the text from the clipboard
    strSelect = Clipboard.GetText
    
    ' Set starting point for reading clipboard text
    intCount = 1
    
    ' Check to see if our initialization string is there
    ' If it's not then we didn't put the text into the clipboard
    ' Let the user know that and exit sub
    If Mid$(strSelect, intCount, Len(App.Title) + 2) <> (Chr(255) & App.Title & Chr(255)) Then
        MsgBox "Invalid clipboard contents!", , "Invalid"
        Exit Sub
    End If
    
    ' Set the starting pount where we search the clipboard
    ' for the list of labels.
    intCount = Len(App.Title) + 2
    
    ' Set the starting area where we will paste the labels
    intPasteStart = lsvLabels.SelectedItem.Index
    
    
    If intPasteStart = 0 Then
        intPasteStart = 1
    End If
    
    ' Loop through the text until we find the seperator
    ' character that seperates the texts from the quantities
    Do
        intCount = intCount + 1                 ' Increment counter
        
        ' If we don't find a text seperator character then add
        ' the currently read character to the temporary text variable
        ' Otherwise we found an entire label text and add the item
        ' to the label list and clear the temporary text variable.
        If Mid$(strSelect, intCount, 1) <> Chr(255) Then
            strTemp = strTemp & Mid$(strSelect, intCount, 1)
        Else
            If lsvLabels.ListItems.Count > 0 Then
                If lsvLabels.SelectedItem.Selected = False Then
                    lsvLabels.ListItems.Add , , strTemp, 23, 23
                Else
                    lsvLabels.ListItems.Add lsvLabels.SelectedItem.Index, , strTemp, 23, 23
                End If
            Else
                lsvLabels.ListItems.Add 1, , strTemp, 23, 23
                lsvLabels.ListItems.Item(lsvLabels.SelectedItem.Index).Selected = False
            End If
            strTemp = ""
        End If
    Loop Until Mid$(strSelect, intCount, 1) = "~"
    
    ' Clear the temporary text variable
    strTemp = ""
    
    ' Loop through the remaining text until we get
    ' to the end of the clipboard text string
    For intCount2 = intCount + 1 To Len(strSelect)
        ' If we don't find a text seperator character then add
        ' the currently read character to the temporary text variable
        ' Otherwise we found an entire label quantity and change
        ' that label's quantity to the value we just read and
        ' clear the temporary text variable.
        If Mid$(strSelect, intCount2, 1) <> Chr(255) Then
            strTemp = strTemp & Mid$(strSelect, intCount2, 1)
        Else
            lsvLabels.ListItems.Item(intPasteStart + intCnt).SubItems(1) = strTemp
            strTemp = ""
            intCnt = intCnt + 1
        End If
    Next intCount2
    
    ' Label list has changed, so allow the user to save it and
    ' Update our status bar and refresh the list
    blnSaved = False
    TotalQty
    mnuRefresh_Click
End Sub

Private Sub mnuPLC_Click()
' This subroutine displays the PLC Sequence dialog box.

    frmPLC.Show 1
End Sub

Private Sub mnuPopupChgQty_Click()
' This subroutine calls the main menu subroutine of the same name.

    mnuChgQty_Click
End Sub

Private Sub mnuPopupCopy_Click()
' This subroutine calls the main menu subroutine of the same name.
    
    mnuCopy_Click
End Sub

Private Sub mnuPopupCut_Click()
' This subroutine calls the main menu subroutine of the same name.
    
    mnuCut_Click
End Sub

Private Sub mnuPopupDelete_Click()
' This subroutine calls the main menu subroutine of the same name.
    
    mnuDelete_Click
End Sub

Private Sub mnuPopupDeleteChar_Click()
' This subroutine calls the main menu subroutine of the same name.
    
    mnuDeleteChar_Click
End Sub

Private Sub mnuPopupEditLabel_Click()
' This subroutine calls the main menu subroutine of the same name.
    
    lsvLabels.StartLabelEdit
End Sub

Private Sub mnuPopupInsert_Click()
' This subroutine calls the main menu subroutine of the same name.
    
    mnuInsertBlank_Click
End Sub

Private Sub mnuPopupInsertchar_Click()
' This subroutine calls the main menu subroutine of the same name.
    
    mnuInsertChar_Click
End Sub

Private Sub mnuPopupPaste_Click()
' This subroutine calls the main menu subroutine of the same name.
    
    mnuPaste_Click
End Sub

Private Sub mnuPopupSelectAll_Click()
' This subroutine calls the main menu subroutine of the same name.
    
    mnuSelectAll_Click
End Sub

Private Sub mnuPrint_Click()
' This subroutine will print the label list in the specified format
    Dim ActualTextHeight As Single

    Dim labLabels() As MyLabel
    Dim strTemp() As String
    Dim intTQty As Long
    Dim sngPrintLeftX As Single
    Dim sngPrintTopY As Single
    Dim itmItem As ListItem
    Dim intResult As Long
    Dim intTemp As Long
    Dim intCnt As Long
    Dim intLPR As Long
    Dim intLPL As Long
    Dim intBlank As Long
    Dim blnTemp As Boolean
    Dim lngLargest As Long
    Dim blnMultLine2 As Boolean
    Dim blnMultLine3 As Boolean
    Dim lngLine As Long
    
    #If DebugMode = 2 Then
        TempText = ""
    #End If
    
    ' If there are no labels in the list, then just exit subroutine.
    If lsvLabels.ListItems.Count = 0 Or IsFormLoaded("frmPrint") Then
        MsgBox "No labels in list or printing in progress already!", vbOKOnly + vbInformation, "Kasa Wire Labels"
        Exit Sub
    End If
    
    If Not CheckIsOKtoPrint(Me.hwnd, (lsvLabels.SelectedItem Is Nothing), blnPrintAll) Then
        MsgBox "Printer was not selected, or not available." & vbCr _
            & "Nothing will be printed.", vbOKOnly
        Exit Sub
    End If
    
    
    
    'FUTURE TODO: INTEGRATE TERMINAL LABELS PROGRAM CODE
'    intTQty = 0
'    intTemp = 0
    
    ' If label format is strip labels, run PrintStrips subroutine instead
'    If intLabelsPerRow = 1 And intLines = 1 Then
'        PrintStrips
'        Exit Sub
'    End If
    
    
    'SNIPPED A
    
        Debug.Print "sngWidth, sngHeight " & sngWidth & ", " & sngHeight
    
    'SNIPPED B
    
    
    ' If the user decides not to print by clicking Cancel, just exit the subroutine
    If Not ConfirmLabelsToPrint(intTQty, blnPrintAll, lsvLabels) Then
        Exit Sub
    End If
    
    
    intTemp = intTQty
    intBlank = 0
    
    ' Add room for blank labels to fill out each row
    Call RoundUpToNearestX(intLabelsPerRow, intTemp)
    If intTemp > intTQty Then
        intBlank = intTemp - intTQty
        intTQty = intTQty + intBlank
    End If
    
    'SNIPPED C
    
    intTQty = intTQty * intCopies
    
    ' Re-Diminsion the array to fit all the labels in the list in the array
    ReDim labLabels(intTQty) As MyLabel
    
    ' Set the Maximum progress bar value to the number of labels to be printed
    prgProgress.Max = intTQty
    
    
    intTemp = 0
    
    ' Set the printers margins by adjusting the numbers
    ' from the currently selected label format
    'SNIPPED D
    sngPrintLeftX = sngLeftMargin - 0.125
    
    ' Adjust the top margin so that no text is cut off
    sngPrintTopY = sngTopMargin - 0.175
    
    
    Call InitPrinter
        ' Set the printed page's height and width
    Printer.Height = sngSpacingTB * 1440
    Printer.Width = intLabelsPerRow * sngSpacingRL * 1440 + sngPrintLeftX * 1440

    Printer.CurrentX = sngPrintLeftX * 1440
    Printer.CurrentY = sngPrintTopY * 1440
    
    #If DebugMode = 2 Then
        PrintTemp "A"
        Call DebugPrinter
    #End If

    If Not AutosizeTallText Then Exit Sub
    
    
    'chg 200508 : Consolidated this routine
    
    
    'Get min / max textwidth size for all labels
    ' Add the label text for each label to the array, depending
    ' on whether the print selection is selected or not.
        For intResult = 1 To intCopies
            
            'Size routine check each label for sizing
            For Each itmItem In lsvLabels.ListItems
                
                'chg 200508
                'If itmItem.Selected = True Then
                If blnPrintAll Or (Not blnPrintAll And itmItem.Selected) Then
                    
                    For intCnt = 1 To Val(itmItem.SubItems(1))
                        'On Error Resume Next
                        strTemp() = Split(itmItem.Text, "|", 3)
                        With labLabels(intTemp)
                            For lngLine = 0 To UBound(strTemp)
                                .strLabel(lngLine) = Trim$(strTemp(lngLine))
                                If Len(.strLabel(lngLine)) = 0 Then
                                    .strLabel(lngLine) = " "
                                    .MaxText = " "
                                Else
                                    If lngLine = 0 Then
                                        'chg 200508
                                        .MaxText = .strLabel(0)
                                    ElseIf lngLine = 1 Then
                                        If Me.TextWidth(.strLabel(1)) > Me.TextWidth(.MaxText) Then
                                            .MaxText = .strLabel(1)
                                        End If
                                        blnMultLine2 = True
                                    ElseIf lngLine = 2 Then
                                        If Me.TextWidth(.strLabel(2)) > Me.TextWidth(.MaxText) Then
                                            .MaxText = .strLabel(2)
                                        End If
                                        blnMultLine3 = True
                                    End If
                                End If
                            Next
                        End With
                        On Error GoTo 0
                        
                        If intAutoSize = 1 Then
                            '200507 : NEW SIZING ROUTINE
                            labLabels(intTemp).lngSize = SizeToText(sngWidth - sngLeftMargin, sngHeight - sngTopMargin, ActualTextHeight, labLabels(intTemp).MaxText)
                            If FeatureMode = WIRE_LABELS Then
                                If labLabels(intTemp).lngSize > 8 Then
                                    labLabels(intTemp).lngSize = 8
                                End If
                            ElseIf FeatureMode = STOCORD_LABELS Then
                                If labLabels(intTemp).lngSize > 18 Then
                                    labLabels(intTemp).lngSize = 18
                                End If
                            End If
                            #If DebugMode = 2 Then
                                PrintTemp "B"
                                PrintTemp "labLabels(intTemp).lngSize = " & labLabels(intTemp).lngSize
                            #End If
                        Else

                            ' Check to see if the width of the text will fit horizontally
                            ' On the printed label, so that no labels overlap the text
                                'CHG 200508 : Using MaxText to determine font size
                                If Printer.TextWidth(labLabels(intTemp).MaxText) > sngWidth * 1440 Then
                                    ' Label width isn't wide enough to fit the text onto the label
                                    ' Notify the user about which label has too long of text
                                    If blnMultLine3 Then
                                        MsgBox "Label " & Chr(34) & Join(labLabels(intTemp).strLabel, "|") & Chr(34) & " is wider than the label!", vbOKOnly + vbInformation, "Label Is Too Wide!"
                                    ElseIf blnMultLine2 Then
                                        MsgBox "Label " & Chr(34) & labLabels(intTemp).strLabel(0) & "|" & labLabels(intTemp).strLabel(1) & Chr(34) & " is wider than the label!", vbOKOnly + vbInformation, "Label Is Too Wide!"
                                    Else
                                        MsgBox "Label " & Chr(34) & labLabels(intTemp).strLabel(0) & Chr(34) & " is wider than the label!", vbOKOnly + vbInformation, "Label Is Too Wide!"
                                    End If
                                    prgProgress.Value = 0
                                    Exit Sub
                                End If

                        End If
                        
                        ' Update progress bar to show current status of loading labels into the array
                        intTemp = intTemp + 1
                        prgProgress.Value = intTemp
                    Next intCnt
                End If
            Next itmItem
            ' Fill the rest of the row with blank labels
            For intCnt = 1 To intBlank
                For lngLine = 0 To 2
                    labLabels(intTemp).strLabel(lngLine) = " "
                Next
                intTemp = intTemp + 1
                prgProgress.Value = intTemp
            Next intCnt
        Next intResult
    
    
    ' Set the largest font size to zero to start with
    lngLargest = 0
    
'-------------------------------------------------

    ' Reset the progress bar value
    prgProgress.Value = 0
    
    ' Show the printer status dialog box
    frmPrint.Show 0, frmMain
    
    ' Update the printer dialog status to 0 labels printed
    frmPrint.lblPrint.Caption = "Printing 0/" & Trim$(Str$(intTQty - (intBlank * intCopies)))
    DoEvents
    
    intTemp = 0
    'chg 200508 : Integrating Optical print routine
    If Not (intOptical = 1) Then
        intCnt = 0
    End If
    
    Do
        ' Set the current vertical position of the printer
        If intOptical = 1 Then
            Printer.CurrentY = sngPrintTopY * 1440
        Else
            ' Set the current vertical position of the printer
            Printer.CurrentY = intCnt * sngSpacingTB * 1440 + sngPrintTopY * 1440
        End If
        #If DebugMode = 2 Then
            PrintTemp "D"
            DebugPrinter
        #End If

        'chg 200508
        '=========
        If blnMultLine3 Or blnMultLine2 Then
            Dim iLast As Integer
            If blnMultLine3 Then
                iLast = 2
            Else
                iLast = 1
            End If
            '==========
            
            ' Print each label the number of lines it needs to be printed
            For intLPL = 0 To iLast
                intLPR = 0

                Do
                    
                    If lngLargest < labLabels(intTemp + intLPR).lngSize Then
                       lngLargest = labLabels(intTemp + intLPR).lngSize
                    End If

                    ' Set the current horizontal position of the printer
                    Printer.CurrentX = intLPR * sngSpacingRL * 1440 + sngPrintLeftX * 1440

                    ' If we are automatically sizing text, then
                    ' set the font size to the size of the current label
                    If intAutoSize = 1 Then
                        If Not labLabels(intTemp + intLPR).lngSize = 0 Then
                            Printer.FontSize = labLabels(intTemp + intLPR).lngSize
                        Else
                            Debug.Assert False
                            'assign minimum size
                            Printer.FontSize = 8
                        End If
                    'else
                    '   Printer font size was set in InitPrinter already.
                    End If

                    ' Print the next label in the list
                    Printer.Print labLabels(intTemp + intLPR).strLabel(intLPL);
                    intLPR = intLPR + 1
                Loop Until (intLPR = intLabelsPerRow) Or ((intTemp + intLPR) = intTQty)

                ' Print a line feed and carriage return to move
                ' the printer to the next line for the next line of text
                If intAutoSize = 1 Then
                    If lngLargest < 8 Then
                        lngLargest = 8
                    End If
                    Printer.FontSize = lngLargest
                End If

                Printer.Print
            Next intLPL


        'End Multiline routine.
        Else

            'Single line per label routine.
            Debug.Assert Not (blnMultLine2 Or blnMultLine3)
            ' Print each label the number of lines it needs to be printed (currently 3)
            For intLPL = 1 To intLines
                intLPR = 0

                lngLargest = 0
                
                Do
                
                    
                    If lngLargest < labLabels(intTemp + intLPR).lngSize Then
                       lngLargest = labLabels(intTemp + intLPR).lngSize
                    End If
                
                    ' Set the current horizontal position of the printer
                    Printer.CurrentX = intLPR * sngSpacingRL * 1440 + sngPrintLeftX * 1440
                    
'                    'CHG 200507 : ALWAYS ADJUSTING THE PRINTER FONT SIZE
                    If intAutoSize = 1 Then
                        If Not labLabels(intTemp + intLPR).lngSize = 0 Then
                            Printer.FontSize = labLabels(intTemp + intLPR).lngSize
                        Else
                            'Any labels with the exception of 'fill in the blank' labels
                            '   at the end of the batch should have had size set already!
                            Debug.Assert Not ((intTQty - (intTemp + intLPR)) > intBlank)
                            'assign minimum size
                            Printer.FontSize = 8
                        End If
                    'else
                    'Printer font size was set in InitPrinter
                    End If
                    

                    #If DebugMode = 2 Then
                        PrintTemp "E"
                        DebugPrinter
                    #End If

                    ' Print the next label in the list
                    Printer.Print labLabels(intTemp + intLPR).strLabel(0);
                    
                    #If DebugMode = 2 Then
                        PrintTemp "F"
                        DebugPrinter
                    #End If
                    
                    intLPR = intLPR + 1
                Loop Until (intLPR = intLabelsPerRow) Or ((intTemp + intLPR) = intTQty)

                
                ' Print a line feed and carriage return to move
                ' the printer to the next line for the next line of text
                If intAutoSize = 1 Then
                    If lngLargest < 8 Then
                        lngLargest = 8
                    End If
                    Debug.Print lngLargest
                    Printer.FontSize = lngLargest
                End If
                Printer.Print
                
                #If DebugMode = 2 Then
                    PrintTemp "G"
                    PrintTemp "intTemp " & intTemp
                    PrintTemp "intLPR " & intLPR
                    PrintTemp "lngLargest " & lngLargest
                    DebugPrinter
                #End If
                
                
            Next intLPL
        End If
'-------------------------------------------------
        
        ' Update the printer status dialog box
        intTemp = intTemp + intLabelsPerRow
        If intTemp > (intTQty - (intBlank * intCopies)) Then
            frmPrint.lblPrint.Caption = "Printing " & Trim$(Str$(intTQty - (intBlank * intCopies))) & " / " & Trim$(Str$(intTQty - (intBlank * intCopies)))
        Else
            frmPrint.lblPrint.Caption = "Printing " & Trim$(Str$(intTemp)) & " / " & Trim$(Str$(intTQty - (intBlank * intCopies)))
        End If
        
        If intOptical = 1 Then
            Printer.NewPage
        Else
            ' Increment counter to adjust for the next row of labels
            intCnt = intCnt + 1
        End If
        DoEvents
    Loop Until intTemp >= intTQty
    
    ' Done printing, finish sending data to the printer

'DBG NOTE:  JUXTAPOSE COMMENTS for the killdoc and enddoc statements
'   in order to test this routine without sending labels to printer

'    Printer.KillDoc
    Printer.EndDoc
        
    ' Update the printer status dialog box to show that we're done printing
    frmPrint.lblPrint.Caption = "Done!"
    DoEvents
    
    ' Wait for 1 second
    Sleep 1000
    
    ' Reset the progress bar
    prgProgress.Value = 0

    ' Close the printer status dialog box
    Unload frmPrint
    Exit Sub
    
ErrorHandler:
    If Err.Number = 396 Or Err.Number = 482 Or Err.Number = 483 _
     Or Err.Number = 484 Then
        '200510 N.F. - Display useful printer error info.
        frmPrintErrorInfo.Show
        frmPrintErrorInfo.DisplayPrinterErrorInfo (Err.Number)
        Err.Clear
    Else                                            ' Any other errors
        MsgBox "Error: " & Err.Number & " - " & Err.Description
    End If
    
    ' Reset the progress bar
    prgProgress.Value = 0
    If IsFormLoaded("frmPrint") Then
        ' Close the printer status dialog box
        Unload frmPrint
    End If
    
End Sub


Private Sub mnuPrintSetup_Click()
' This subroutine calls the built-in windows printer setup.

    On Error GoTo ErrorHandler
    
    frmPrinterSetup.Show 1
    'If lsvLabels.SelectedItem Is Nothing Then
    '    SelectPrinter Me.hwnd, True, intCopies, True
    'Else
    '    SelectPrinter Me.hwnd, True, intCopies, False
    'End If
    
    Exit Sub                                        ' Bypass Error Handler

ErrorHandler:
    If Err.Number = 32755 Then                      ' If Cancel is Selected
        Exit Sub
    ElseIf Err.Number = 396 Or Err.Number = 482 Or Err.Number = 483 _
     Or Err.Number = 484 Then
        '200510 N.F. - Display useful printer error info.
        frmPrintErrorInfo.Show
        frmPrintErrorInfo.DisplayPrinterErrorInfo (Err.Number)
        Err.Clear
    Else                                            ' Any other errors
        MsgBox "Error: " & Err.Number & " - " & Err.Description
    End If
End Sub

Private Sub mnuProgressBar_Click()
' This subroutine changes the visibility status of the progress bar.
    
    prgProgress.Visible = Not prgProgress.Visible   ' Toggle Progress Bar Visibility
    CheckVisualStatus                               ' Update checked menus
End Sub

Private Sub mnuRefresh_Click()
' This subroutine refreshes the list of labels.
    
    lsvLabels.Refresh
End Sub

Private Sub mnuSave_Click()
' This subroutine allows the user to save the current list.
' It also puts a default filename into the box.
' The filename is either the current filename, or Untitled.

    On Error GoTo ErrorHandler
    
    Dim itmItem As ListItem
    
    ' If we don't have any items, just exit sub.
    If lsvLabels.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    ' Set the dialog box filter to *.LBL and *.* files, then
    ' set the dialog box filename to the default name, and
    ' open the dialog box
    cdlDialog.Filter = "Label List Files (*.lbl)|*.lbl|All Files (*.*)|*.*"
    cdlDialog.Flags = cdlOFNHideReadOnly
    cdlDialog.FileName = strFileName
    cdlDialog.ShowSave
    
    ' If the user didn't select a name to save it as just exit sub.
    If cdlDialog.FileName = "" Then
        Exit Sub
    End If
    
    ' Add the file to the documents menu on the start menu
    SHAddToRecentDocs SHARD_PATH, cdlDialog.FileName
    
    ' Set the default filename to the name they just selected.
    strFileName = cdlDialog.FileName
    
    ' Set the progress bar value to 0 and the maximum to the number
    ' of labels we have in our list.
    prgProgress.Value = 0
    prgProgress.Max = lsvLabels.ListItems.Count
    
    ' Set the mouse pointer to the hourglass
    frmMain.MousePointer = 11
    
    ' Open the file for output
    Open cdlDialog.FileName For Output As #1
        
    ' Print the data for each label to the open file
    ' Update the progress bar to show the status of saving
    For Each itmItem In lsvLabels.ListItems
        Print #1, itmItem.Text
        Print #1, itmItem.SubItems(1)
        prgProgress.Value = prgProgress.Value + 1
        DoEvents
    Next
    
    ' Close the file
    Close #1
    
    ' We just saved it, no need to save again
    blnSaved = True
    
    ' Set the mouse pointer back to the default
    frmMain.MousePointer = 0
    
    ' Set the progress bar value to 0
    prgProgress.Value = 0
    
    ' Update the form's caption to the new filename
    Me.Caption = GetFormCaption(Me.Name) & " - " & strFileName
    
    Exit Sub
ErrorHandler:
    If Err.Number = 32755 Then                      ' If Cancel is Selected
        Exit Sub
    Else                                            ' Any other errors
        MsgBox "Error: " & Err.Number & " - " & Err.Description
    End If
End Sub

Private Sub mnuSelectAll_Click()
' This subroutine selects all the labels in the list

    Dim itmItem As ListItem
    
    ' If there are no labels in the list just exit this subroutine
    If lsvLabels.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    ' Cycle through every label in the list and select it
    For Each itmItem In lsvLabels.ListItems
        itmItem.Selected = True
    Next
End Sub

Private Sub mnuSLC_Click()
' This subroutine shows the SLC Sequence dialog box

    frmSLC.Show 1
End Sub

Private Sub mnuSort_Click()
' This subroutine sorts the label text in ascending order.

    With lsvLabels
        .SortKey = 0                                ' Sort on label text
        .SortOrder = lvwAscending                   ' Sort in ascending order
        .Sorted = True                              ' Sort the list
        .Sorted = False                             ' Prevent sorting new items
    End With
End Sub

Private Sub mnuStatusBar_Click()
' This subroutine toggles the visibility of the status bar.

    stbStatus.Visible = Not stbStatus.Visible       ' Toggle Status Bar Visibility
    CheckVisualStatus                               ' Update checked menus
End Sub

Private Sub mnuToolbar_Click()
' This subroutine toggles the visibility of the toolbar.

    tlbToolbar.Visible = Not tlbToolbar.Visible     ' Toggle Toolbar Visibility
    CheckVisualStatus                               ' Update checked menus
End Sub

Private Sub tlbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
' This subroutine checks to see which toolbar button was
' clicked and executes the corresponding main menu command
' that is associated with that button.

    Select Case Button.Key
        Case "New"                  ' They clicked New Button
            ' Act as if the user clicked the appropriate menu instead
            mnuNew_Click
        
        Case "Open"                 ' They clicked Open Button
            ' Act as if the user clicked the appropriate menu instead
            mnuOpen_Click
            
        Case "Import"               ' They clicked Import Button
            ' Act as if the user clicked the appropriate menu instead
            mnuImport_Click
            
        Case "Export"               ' They clicked Export Button
            ' Act as if the user clicked the appropriate menu instead
            mnuExport_Click
            
        Case "Save"                 ' They clicked Save Button
            ' Act as if the user clicked the appropriate menu instead
            mnuSave_Click
            
        Case "Format"               ' They clicked Format Button
            ' Act as if the user clicked the appropriate menu instead
            mnuLabelFormat_Click
            
        Case "Print"                ' They clicked Print Button
            ' Act as if the user clicked the appropriate menu instead
            mnuPrint_Click
            
        Case "DeleteLabel"          ' They clicked DeleteLabel Button
            ' Act as if the user clicked the appropriate menu instead
            mnuDelete_Click
            
        Case "Cut"                  ' They clicked Cut Button
            ' Act as if the user clicked the appropriate menu instead
            mnuCut_Click
        
        Case "Copy"                 ' They clicked Copy Button
            ' Act as if the user clicked the appropriate menu instead
            mnuCopy_Click
        
        Case "Paste"                ' They clicked Paste Button
            ' Act as if the user clicked the appropriate menu instead
            mnuPaste_Click
            
        Case "Font"                 ' They clicked Font Button
            ' Act as if the user clicked the appropriate menu instead
            mnuFont_Click
        
        Case "Find"                 ' They clicked Find Button
            ' Act as if the user clicked the appropriate menu instead
            mnuFindNext_Click
            
        Case "Sort"                 ' They clicked Sort Button
            ' Act as if the user clicked the appropriate menu instead
            mnuSort_Click
            
        Case "InsertChar"           ' They clicked InsertChar Button
            ' Act as if the user clicked the appropriate menu instead
            mnuInsertChar_Click
            
        Case "DeleteChar"           ' They clicked DeleteChar Button
            ' Act as if the user clicked the appropriate menu instead
            mnuDeleteChar_Click
            
        Case "Quantity"             ' They clicked Quantity Button
            ' Act as if the user clicked the appropriate menu instead
            mnuChgQty_Click
            
        Case "Decimal"              ' They clicked Decimal Button
            ' Act as if the user clicked the appropriate menu instead
            mnuDecimal_Click
        
        Case "Alpha"                ' They clicked Alpha Button
            ' Act as if the user clicked the appropriate menu instead
            mnuAlpha_Click
            
        Case "PLC"                  ' They clicked PLC Button
            ' Act as if the user clicked the appropriate menu instead
            mnuPLC_Click
            
        Case "SLC"                  ' They clicked SLC Button
            ' Act as if the user clicked the appropriate menu instead
            mnuSLC_Click
        
        Case "Help"                 ' They clicked Help Button
            ' Act as if the user clicked the appropriate menu instead
            mnuContents_Click
            
        Case Else                   ' They clicked some other button we don't know about?
            MsgBox "Select another button please!", , "Error"
            
    End Select
End Sub

Private Sub txtLabelQty_Change()
' This subroutine updates the status bar with what is currently
' being typed in the new label box.  It also makes sure that there
' is a valid quantity in the text box before allowing the user
' to accept the entry as a new label.

    ' Update the status bar with the current new label data
    stbStatus.Panels(1).Text = "New Label"
    stbStatus.Panels(2).Text = "Label Text: " & Trim$(txtLabelText.Text)
    stbStatus.Panels(3).Text = "Qty: " & Trim$(Val(txtLabelQty.Text))

    ' If there is no value for the quantity of labels, don't
    ' let the user click the add button
    If Val(txtLabelQty.Text) < 1 Then
        cmdAdd.Enabled = False
        txtLabelQty.Text = "0"
        Exit Sub
    Else
        cmdAdd.Enabled = True
    End If
    
    ' Prevent the user from pasting anything but a number
    txtLabelQty.Text = Trim$(Str$(Val(txtLabelQty.Text)))
    txtLabelQty.SelStart = Len(txtLabelQty.Text)
End Sub

Private Sub txtLabelQty_GotFocus()
' This subroutine disables the cut, copy, paste, and delete
' commands and lets the user know if the label text already
' exists so that they do not type in two sets of identical
' labels by accident.

    Dim intCount As Long
    
    ' Don't let the user try to copy or paste a label
    ' when there isn't a label selected
    mnuCopy.Enabled = False
    tlbToolbar.Buttons(12).Enabled = False
    mnuCut.Enabled = False
    tlbToolbar.Buttons(11).Enabled = False
    mnuDelete.Enabled = False
    tlbToolbar.Buttons(10).Enabled = False
    mnuPaste.Enabled = False
    mnuPopupPaste.Enabled = False
    tlbToolbar.Buttons(13).Enabled = False

    ' Highlight the text in the text box so we don't need to backspace
    SelectText
    
    If lsvLabels.ListItems.Count = 0 Then
        Exit Sub                                ' If There Are No Items Exit Subroutine
    End If
    
    ' Check List To See If Label Already Exists
    ' If It Does, Notify User That It Already Exists
    For intCount = 1 To lsvLabels.ListItems.Count
        If lsvLabels.ListItems.Item(intCount).Text = txtLabelText.Text Then
            lblUserMsg.Caption = "* Label already exists.  Another label with same text will be created."
        End If
    Next intCount
End Sub

Private Sub txtLabelQty_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it adds the current label text and quantity
' to the label list as a new label, and puts the user in the
' label text field so that they may enter the next label.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete _
    And KeyAscii <> vbKeyLeft And KeyAscii <> vbKeyRight _
    And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyTab _
    And KeyAscii <> vbKeyDecimal Then
        KeyAscii = 0
    End If
    If KeyAscii = vbKeyReturn And Val(txtLabelQty.Text) > 0 Then
        KeyAscii = 0                                ' Void out the enter key that was pressed
        cmdAdd_Click                                ' Add label to the list
    End If
End Sub

Private Sub txtLabelText_Change()
' This subroutine updates the status bar with the current
' text entered by the user as the new label.

    stbStatus.Panels(1).Text = "New Label"
    stbStatus.Panels(2).Text = "Label Text: " & Trim$(txtLabelText.Text)
    stbStatus.Panels(3).Text = "Qty: " & Trim$(Val(txtLabelQty.Text))
End Sub

Private Sub txtLabelText_GotFocus()
' This subroutine disables the copy, cut, paste, and delete
' commands throughout the program, so that they cannot cause
' the program to freeze up with invalid data.  It also updates
' the status bar with the information of the last label entered.

    ' Disable the cut, copy, paste menus, because they're for the labels only
    mnuCopy.Enabled = False
    tlbToolbar.Buttons(12).Enabled = False
    mnuCut.Enabled = False
    tlbToolbar.Buttons(11).Enabled = False
    mnuDelete.Enabled = False
    tlbToolbar.Buttons(10).Enabled = False
    mnuPaste.Enabled = False
    mnuPopupPaste.Enabled = False
    tlbToolbar.Buttons(13).Enabled = False
    
    ' Deselect all the labels in the list
    UnSelect
    
    ' Update the status bar
    If Trim$(txtLabelText.Text) = "" And lsvLabels.ListItems.Count > 0 Then
        stbStatus.Panels.Item(1).Text = "Label: " & lsvLabels.SelectedItem.Index & "/" & lsvLabels.ListItems.Count
        stbStatus.Panels.Item(2).Text = "Label Text: " & lsvLabels.SelectedItem.Text
        stbStatus.Panels.Item(3).Text = "Qty: " & Trim$(lsvLabels.SelectedItem.SubItems(1))
    End If
    SelectText                                  ' Select all text in text box
End Sub

Private Sub txtLabelText_KeyPress(KeyAscii As Integer)
' This subroutine checks to see if the Enter key was pressed.
' If it was, cancel the enter key, and send a tab key instead
' to make the focus shift to the label quantity box instead.

    If KeyAscii = vbKeyReturn Then
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
        KeyAscii = 0
    End If
End Sub

Private Sub txtLabelText_LostFocus()
' This subroutine strips off any whitespace or null characters
' off of the label text as the text box loses focus.

    txtLabelText.Text = Trim$(txtLabelText.Text)
End Sub

Private Sub UnSelect()
' This subroutine Deselects all the items in the list
    Dim itmItem As ListItem
    
    For Each itmItem In lsvLabels.ListItems
        itmItem.Selected = False
    Next itmItem
End Sub






'chg 200505
Public Function SizeToText(ByRef MaxTextWidth As Single, _
 ByRef MaxTextHeight As Single, ByRef ActualTextHeight As Single, _
 ByVal s As String, Optional ByVal s2 As String = "") As Single
    
    Dim bChanged As Boolean
    Dim OrigSize As Single
    Dim LastFontSize As Single
    
    
    If s = "" Then
        Debug.Assert False
        SizeToText = 0
        Exit Function
    End If
    'chg 46
    'Font size minimum is as if text is at least 4 characters
    bChanged = False
    Do While Len(s) < 4
        bChanged = True
        s = s & s
    Loop
    If bChanged Then
        s = Left(s, 4)
    End If
    '==
    
    With Me
        Font.Bold = cdlDialog.FontBold
        Font.Italic = cdlDialog.FontItalic
        Font.Name = cdlDialog.FontName
        Font.Strikethrough = cdlDialog.FontStrikethru
        Font.Underline = cdlDialog.FontUnderline
    End With
    
    OrigSize = Me.Font.Size
    
    Dim StartSize As Single
    Dim RetVal As Single
    
    StartSize = 3.8
    Me.Font.Size = StartSize
    
    'Size to width
    If s2 <> "" Then
        'making both lines same font size
        Do
            If LastFontSize < Me.Font.Size Then
                LastFontSize = Me.Font.Size
            End If
            StartSize = StartSize + 0.2
            Me.Font.Size = StartSize
        Loop Until Me.TextWidth(s) >= (MaxTextWidth * 1440) Or _
         Me.TextWidth(s2) >= (MaxTextWidth * 1440)
    
    Else
        'only need to size first line
        Do
            If LastFontSize < Me.Font.Size Then
                LastFontSize = Me.Font.Size
            End If
            StartSize = StartSize + 0.2
            Me.Font.Size = StartSize
        Loop Until Me.TextWidth(s) >= (MaxTextWidth * 1440)
    End If
    
    'This was previous way to limit height, but it doesn't work
    'for variable-spaced terminals
'    If Me.Font.Size > 10 Then
'        RetVal = 10
'    Else
'        RetVal = LastFontSize
'    End If
    
    RetVal = LastFontSize
    Debug.Print "Sizing to Width: FontSize = " & RetVal
    
    'Back down font size to size the proper height
    If s2 <> "" Then
        'making both lines same font size
        Do While Me.TextHeight(s) + Me.TextHeight(s2) > (MaxTextHeight * 1440)
            If LastFontSize > Me.Font.Size Then
                LastFontSize = Me.Font.Size
            End If
            StartSize = StartSize - 0.2
            Me.Font.Size = StartSize
        Loop
        
        'Set return value for actual text height
        ActualTextHeight = Me.TextHeight(s) + Me.TextHeight(s2)
    Else
    
        Do While Me.TextHeight(s) > (MaxTextHeight * 1440)
            If LastFontSize > Me.Font.Size Then
                LastFontSize = Me.Font.Size
            End If
            StartSize = StartSize - 0.2
            Me.Font.Size = StartSize
        Loop
    
        'Set return value for actual text height
        ActualTextHeight = Me.TextHeight(s)
    End If
    
    RetVal = LastFontSize
    Debug.Print "Fit to Height: FontSize = " & RetVal
    SizeToText = RetVal
    
    Me.Font.Size = OrigSize
    
End Function




Private Sub OpenCommandLine()
' This subroutine allows the user to open a label list file
' from the command line that they have previously saved.

    On Error GoTo ErrorHandler
    
    Dim intCount As Long
    Dim Result As Long
    Dim strFileTemp As String
    Dim strTemp1 As String
    Dim strTemp2 As String
    
    strFileTemp = Mid$(Trim$(Command$), 2, Len(Trim$(Command$)) - 2)
    
    ' Set the default filename to the file they just opened
    strFileName = strFileTemp
    
    ' Set the mouse pointer to an hourglass
    frmMain.MousePointer = 11
    
    ' Clear the label list
    lsvLabels.ListItems.Clear
    
    ' Open the file for input
    Open strFileTemp For Input As #1
    
    ' Add each item to the list from the data in the file
    ' and keep reading them in until we get to the end of
    ' the file
    While Not EOF(1)
        intCount = intCount + 1
        Input #1, strTemp1, strTemp2
        If Trim$(strTemp2) <> "" Then
            lsvLabels.ListItems.Add , , Trim$(strTemp1), 23, 23
            lsvLabels.ListItems(intCount).SubItems(1) = Trim$(strTemp2)
        Else
            intCount = intCount - 1
        End If
    Wend
    
    ' Close the file
    Close #1
        
    If Not blnLockFormat Then
        ' Ask the user if they want to change the current label format
        Result = MsgBox("The current label format is: " & Chr(34) & Trim$(strLabelFormat) & Chr(34) & "." & vbCrLf & "Do you want to change this?", vbYesNo + vbExclamation, "Change Format?")
        If Result = vbYes Then
            frmFormat.Show 1
        End If
    End If
    
    ' Unselect the first item that is selected by default
    lsvLabels.ListItems(1).Selected = False
    
    ' Set the mouse pointer back to the default
    frmMain.MousePointer = 0
    
    ' Clear the search string text
    strFind = ""

    ' File is the same as when it was opened, so no need to save
    blnSaved = True
    
    ' Re-Total the label quantites and update the status bar
    TotalQty
    
    ' Add the file to the documents menu on the start menu
    SHAddToRecentDocs SHARD_PATH, strFileTemp
    
    
    Exit Sub                                        ' Bypass Error Handler

ErrorHandler:
    If Err.Number = 32755 Then                      ' If Cancel is Selected
        Exit Sub
    Else                                            ' Any other errors
        MsgBox "Error: " & Err.Number & " - " & Err.Description
    End If
End Sub


'===============================
'   BEGIN DEBUGMODE CODE
'===============================
#If DebugMode = 1 Then
    Private Sub PrintStripsPreview()
    
        ' Display the printer status dialog and update it's text
        frmPrintPreview.Show 0, frmMain
    
    
    End Sub

    Private Sub mnuPrintPreview_Click()
        Call PrintStripsPreview
    
    End Sub

#End If

#If DebugMode = 2 Then

    Private Sub DebugPrinter()
    
        With Printer
            PrintTemp .CurrentX
            PrintTemp .CurrentY
            PrintTemp .FontBold
            PrintTemp .FontItalic
            PrintTemp .FontName
            PrintTemp .FontSize
            PrintTemp .FontStrikethru
            PrintTemp .FontUnderline
            PrintTemp .Height
            PrintTemp .Width
        End With
    
    End Sub

    Private Sub PrintTemp(ByVal s As String)
    
    If intOptical = 1 Then
        TempText2 = TempText2 & s & vbCrLf
    Else
        TempText = TempText & s & vbCrLf
    End If
    
    End Sub

#End If
'===============================
'   END DEBUGMODE CODE
'===============================



'===============================
'   BEGIN COPIED CODE
'===============================
'NOTE:  GET CURRENT PRINTSTRIPS ROUTINE FROM KASA TERMINAL LABELS SOURCE IN CASE
'   THIS ISN'T THE LATEST
'
'chg 200504
Private Sub PrintStrips()
' This subroutine will print the label list onto the strip
' labels instead of the normal labels, by using the format
' specified through the label format dialog box.
    Dim NextY As Single
    Dim ActualTextHeight As Single
    
    Dim sFirstLine As String
    Dim sSecondLine As String
    Dim PageHeight As Single
    Dim CurLabel As Long
    Dim PgLastLabel As Long
    Dim Pages As Integer
    Dim CurPage As Integer
    Dim strLabel() As String
    Dim lngSize() As Long
    Dim intTQty As Long
    Dim sngPrintLeftX As Single
    Dim sngPrintTopY As Single
    Dim itmItem As ListItem
    Dim intResult As Long
    Dim intTemp As Long
    Dim intCnt As Long
    Dim intLPR As Long
    Dim intCut As Long
    Dim intPos As Long
    Dim intMax As Long
    Dim intPageOffset As Long
    Dim blnTemp As Boolean
    Dim lngMinSize As Long
    
    If FeatureMode <> TERMINAL_STRIPS Then
        Debug.Assert False
        Exit Sub
    End If

    intTQty = 0
    intTemp = 0
    
    'chg 200504 - SplitText routine no longer adds 1 to this var
    'Max chars per printed line
    intMax = 7
    'Max Page height is 49.2", using 49" (49 * 1440) Twips
    PageHeight = 49
    
    If SelectForm(App.Title, Me.hwnd, 4, 49) = 0 Then
        ' Selection failed!
        MsgBox "Cannot Print: '" & App.Title & "' form could not be set.", vbOKOnly + vbCritical, App.Title
        Exit Sub
    End If
    
    ' If we are only printing the selected items, the only count
    ' the quantities of the selected labels.
    If Not blnPrintAll Then
        ' Only count the selected labels
        For Each itmItem In lsvLabels.ListItems
            If itmItem.Selected = True Then
                intTQty = intTQty + Val(itmItem.SubItems(1))
            End If
            intTemp = intTemp + Val(itmItem.SubItems(1))
        Next itmItem
        ' Ask user if they really want to print these labels
        If intCopies > 1 Then
            intResult = MsgBox(Trim$(Str$(intTQty)) & " / " & Trim$(Str$(intTemp)) & " labels to print in " & Chr(34) & strLabelFormat & Chr(34) & " format." & vbCrLf & "This will print " & intCopies & " copies." & vbCrLf & "Using " & Printer.DeviceName & " on port " & Printer.Port & vbCrLf & vbCrLf & "Align label sheet in printer and click Ok.", vbOKCancel + vbInformation, "Print")
        Else
            intResult = MsgBox(Trim$(Str$(intTQty)) & " / " & Trim$(Str$(intTemp)) & " labels to print in " & Chr(34) & strLabelFormat & Chr(34) & " format." & vbCrLf & "This will print " & intCopies & " copy." & vbCrLf & "Using " & Printer.DeviceName & " on port " & Printer.Port & vbCrLf & vbCrLf & "Align label sheet in printer and click Ok.", vbOKCancel + vbInformation, "Print")
        End If
    Else
        ' Count all the labels in the list
        For Each itmItem In lsvLabels.ListItems
            intTQty = intTQty + Val(itmItem.SubItems(1))
        Next itmItem
        ' Ask user if they really want to print these labels
        If intCopies > 1 Then
            intResult = MsgBox(Trim$(Str$(intTQty)) & " labels to print in " & Chr(34) & strLabelFormat & Chr(34) & " format." & vbCrLf & "This will print " & intCopies & " copies." & vbCrLf & "Using " & Printer.DeviceName & " on port " & Printer.Port & vbCrLf & vbCrLf & "Align label sheet in printer and click Ok.", vbOKCancel + vbInformation, "Print")
        Else
            intResult = MsgBox(Trim$(Str$(intTQty)) & " labels to print in " & Chr(34) & strLabelFormat & Chr(34) & " format." & vbCrLf & "This will print " & intCopies & " copy." & vbCrLf & "Using " & Printer.DeviceName & " on port " & Printer.Port & vbCrLf & vbCrLf & "Align label sheet in printer and click Ok.", vbOKCancel + vbInformation, "Print")
        End If
    End If
    
    ' If the user cancels the print, just exit the subroutine
    If intResult = vbCancel Then
        Exit Sub
    End If
    
    ' Recalculate the total quantity of labels by multiplying
    ' the single list times the number of copies needed
    intTQty = intTQty * intCopies
    
    'chg 200504
    'DOUBLE CHECK QUANTITY to ensure non-zero!
    If intTQty = 0 Then
        MsgBox "No Labels Selected.  Nothing to do!", vbExclamation
        Exit Sub
    End If
    
    ' Re-Dimension the array to hold all the labels it counted
    ReDim strLabel(intTQty - 1)
    ReDim lngSize(intTQty - 1)
    
    
    
    ' Display the printer status dialog and update it's text
    frmPrint.Show 0, frmMain
    frmPrint.lblPrint.Caption = "Printing 1/" & Trim$(Str$(intTQty))
    DoEvents
    
    
    'Add a dummy labels to provide space at end to cut the labels from spool
    intTQty = intTQty + 3
    ReDim Preserve strLabel(intTQty - 1)
    ReDim Preserve lngSize(intTQty - 1)
    strLabel(intTQty - 3) = "eoj marker"
    strLabel(intTQty - 2) = "eoj marker"
    strLabel(intTQty - 1) = "eoj marker"
    
    intTemp = 0
    
    'Split the text on the wire labels to two lines if necessary
    If Not blnPrintAll Then
        For intResult = 1 To intCopies
            ' Add only the selected label texts to the array
            For Each itmItem In lsvLabels.ListItems
                If itmItem.Selected = True Then
                    For intCnt = 1 To Val(itmItem.SubItems(1))
                        strLabel(intTemp) = SplitText(itmItem.Text, intMax)
                        intTemp = intTemp + 1
'                        prgProgress.Value = intTemp
                    Next intCnt
                End If
            Next itmItem
        Next intResult
    Else
        For intResult = 1 To intCopies
            ' Add all the label texts to the array
            For Each itmItem In lsvLabels.ListItems
                For intCnt = 1 To Val(itmItem.SubItems(1))
                    strLabel(intTemp) = SplitText(itmItem.Text, intMax)
                    intTemp = intTemp + 1
'                    prgProgress.Value = intTemp
                Next intCnt
            Next itmItem
        Next intResult
    End If
    
    
    prgProgress.Min = LBound(strLabel)
    prgProgress.Max = UBound(strLabel)
    
    intTemp = 0
    
    ' Adjust the printer's margins from what the label format is
    sngPrintLeftX = sngLeftMargin - 0.125
    
    'Strip labels ignore top margin
    sngPrintTopY = 0
'    sngPrintTopY = sngTopMargin - 0.175
    
    ' Set the printer's font settings
    With Printer
        .FontBold = cdlDialog.FontBold
        .FontItalic = cdlDialog.FontItalic
        .FontName = cdlDialog.FontName
        If intAutoSize = 1 Then
            'do nothing
        Else
            .FontSize = cdlDialog.FontSize
        End If
        .FontStrikethru = cdlDialog.FontStrikethru
        .FontUnderline = cdlDialog.FontUnderline
    End With
    
  
    'chg 200504
    'Print until all 'pages' are finished
    
    Printer.Width = ((intLabelsPerRow - 1) * sngSpacingRL + sngWidth + sngLeftMargin) * 1440
    Debug.Print "Printer width : " & Printer.Width / Screen.TwipsPerPixelX / 96
    
    'Set number of pages
    If ((intTQty) * sngSpacingTB) > (PageHeight - sngPrintTopY) Then
        Pages = CInt(((intTQty) * sngSpacingTB) / (PageHeight - sngPrintTopY)) + 1
        If Pages - ((intTQty) * sngSpacingTB) / (PageHeight - sngPrintTopY) > 1 Then
            Pages = Pages - 1
        End If
    Else
        Pages = 1
    End If
    
    CurPage = 1
    CurLabel = 0
    
    
    Do
    
    
    On Error Resume Next
    
    If CurPage = Pages Then
   
        PgLastLabel = UBound(strLabel)
    
        Printer.Height = sngPrintTopY * 1440 + (PgLastLabel - CurLabel + 1) * sngSpacingTB * 1440
    
    Else
        

        Debug.Assert (CurPage < Pages)
    
        PgLastLabel = CurLabel + CInt((PageHeight - sngPrintTopY) / sngSpacingTB) - 1
    
        Printer.Height = 1440 * PageHeight
    
    End If
    
    'Debug.Print "Printer height: " & Printer.Height / Screen.TwipsPerPixelX / 96
    
    If Err Then
        'bmk todo: comment the stop
        Debug.Print Err.Number & " " & Err.Description
        MsgBox "Unexpected Error: " & vbCr & Err.Number & vbCr & Err.Description, "PrintStrips"
        'Stop
        Err.Clear
    End If
    
    'After the labels have been chosen, text has been split if necessary,
    '   then lines have been sized to proper font, we're finally ready to print!
    
    intPageOffset = CurLabel
    
    For intCnt = CurLabel To PgLastLabel
        
'        ' Set the current printer position
'        If strLabel(intCnt) <> "eoj marker" Then
'            Printer.CurrentY = (intCnt - intPageOffset) * sngSpacingTB * 1440 + sngPrintTopY * 1440
'        Else
'            Printer.CurrentY = (intCnt - intPageOffset) * sngSpacingTB * 1440 + sngPrintTopY * 1440 + 1440 * 0.125
'        End If
        
        
        ' Loop through all lines of text of the label
        Do
            ' Set the horizontal position of the printer
            Printer.CurrentX = sngPrintLeftX * 1440
            ' Check the text for a return and line feed
            If InStr(1, strLabel(intCnt), vbCrLf) = 0 Then
                ' There is no line break representing
                '   two lines of text, so just print the text.
                sFirstLine = strLabel(intCnt)
                sSecondLine = ""
                strLabel(intCnt) = ""
            Else
                ' Print all the text up to but not including the line break
                ' Then trim off remaining text w/o line break
                sFirstLine = Left$(strLabel(intCnt), InStr(1, strLabel(intCnt), vbCrLf) - 1)
                sSecondLine = Mid$(strLabel(intCnt), InStr(1, strLabel(intCnt), vbCrLf) + Len(vbCrLf))
                'strLabel(intCnt) = Mid$(strLabel(intCnt), InStr(1, strLabel(intCnt), vbCrLf) + Len(vbCrLf))
                strLabel(intCnt) = ""
            End If
        
        
            ' If we are automatically sizing the text, then set
            ' the printer font size to the size of the text for
            ' this specific label
            If intAutoSize = 1 Then
                If Trim(sFirstLine) <> "" Then
                    If Len(sFirstLine) = 4 Or Len(sFirstLine) = 5 Then
                        'some reason this many chars spills over the allotted width
                        lngSize(intCnt) = SizeToText(sngWidth - 1.5 / 16, sngHeight - 1 / 16, ActualTextHeight, sFirstLine, sSecondLine)
                    Else
                        lngSize(intCnt) = SizeToText(sngWidth - 1 / 16, sngHeight - 1 / 16, ActualTextHeight, sFirstLine, sSecondLine)
                    End If
                Else
                    sFirstLine = " "
                End If
            
                If lngSize(intCnt) >= 3 Then
                    Printer.Font.Size = lngSize(intCnt)
                Else
                    Debug.Assert False
                    Printer.Font.Size = 3
                End If
            
            End If
            
            
        ' Set the current printer position
        If strLabel(intCnt) <> "eoj marker" Then
            'chg 20050505
            'The labels should be centered vertically rather than top justified
            Printer.CurrentY = (intCnt - intPageOffset) * sngSpacingTB * 1440 + sngPrintTopY * 1440 + (sngHeight * 1440 - ActualTextHeight) * 0.5
        Else
            Printer.CurrentY = (intCnt - intPageOffset) * sngSpacingTB * 1440 + sngPrintTopY * 1440 + (sngHeight * 1440 - ActualTextHeight) * 0.5 + 1440 * 0.125
        End If
            
            Printer.Print sFirstLine
            If sSecondLine <> "" Then
                ' Set the horizontal position of the printer
                Printer.CurrentX = sngPrintLeftX * 1440
                Printer.Print sSecondLine
            End If
            
        ' Continue looping until the entire label is printed
        Loop Until Len(strLabel(intCnt)) = 0
        
        
        '20050505
        If intCnt < PgLastLabel Then
            If strLabel(intCnt + 1) <> "eoj marker" Then
                NextY = (intCnt + 1 - intPageOffset) * sngSpacingTB * 1440 + sngPrintTopY * 1440
            Else
                NextY = (intCnt + 1 - intPageOffset) * sngSpacingTB * 1440 + sngPrintTopY * 1440 + 1440 * 0.125
            End If
        Else
            NextY = Printer.CurrentY
        End If
        
        'Safety valve
        'shouldn't happen if we set PgLastLabel correctly!!
        If NextY > Printer.Height Then
            Debug.Assert (intCnt = PgLastLabel)
            Exit For
        End If
        
        ' Update the progress bar
        'prgProgress.Value = intCnt + LBound(strLabel) + 1
        prgProgress.Value = intCnt
        
        'frmPrint.lblPrint.Caption = "Printing " & Trim$(Str$(intCnt + LBound(strLabel) + 1)) & " / " & Trim$(Str$(intTQty))
        frmPrint.lblPrint.Caption = "Printing " & Trim$(Str$(intCnt + 1)) & " / " & Trim$(Str$(intTQty))
        
        DoEvents
    Next 'intCnt
    
    If Pages > 1 And CurPage < Pages Then
            Printer.NewPage
            
        'Further new pages aren't really new pages, it's just a strip label,
        '   so we can remove the top margin
'        sngPrintTopY = 0
    End If
    
    CurPage = CurPage + 1
    
    CurLabel = PgLastLabel + 1
    
    Loop Until CurLabel > UBound(strLabel)
    
    If Pages = 1 Or (Pages > 1 And CurPage > Pages) Then
        Printer.Print " "
        Printer.Print " "
        Printer.Print " "
        
        ' Finish sending data to the printer, we're done printing
        Printer.EndDoc
    End If
    
    ' Let user know we're done printing.
    frmPrint.lblPrint.Caption = "Done!"
    DoEvents

    ' Wait 1 second before closing the printer status
    Sleep 1000
    prgProgress.Value = 0
    
    ' Close the printer status dialog box
    Unload frmPrint
End Sub
'===============================
'   END COPIED CODE
'===============================

