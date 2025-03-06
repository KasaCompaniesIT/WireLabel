VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrinterSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Printer Setup"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCopies 
      Caption         =   "Copies"
      Height          =   1215
      Left            =   2760
      TabIndex        =   12
      Top             =   1560
      Width           =   2415
      Begin MSComCtl2.UpDown updCopies 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCopies"
         BuddyDispid     =   196620
         OrigLeft        =   1800
         OrigTop         =   480
         OrigRight       =   2040
         OrigBottom      =   855
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCopies 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "1"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblCopies 
         Caption         =   "Co&pies: "
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame fraRange 
      Caption         =   "Print Range"
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   2415
      Begin VB.OptionButton optSelection 
         Caption         =   "Se&lection"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optAll 
         Caption         =   "&All"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame fraPrinterSetup 
      Caption         =   "Printer"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton cmdSetup 
         Caption         =   "Printer &Setup"
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboPrinter 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblName 
         Caption         =   "&Name: "
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmPrinterSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function PrinterProperties Lib "winspool.drv" _
  (ByVal hwnd As Long, ByVal hPrinter As Long) As Long

Private Declare Function OpenPrinter Lib "winspool.drv" _
  Alias "OpenPrinterA" (ByVal pPrinterName As String, _
  phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long

Private Declare Function ClosePrinter Lib "winspool.drv" _
  (ByVal hPrinter As Long) As Long
  
Private Type PRINTER_DEFAULTS
   pDatatype As Long ' String
   pDevMode As Long
   pDesiredAccess As Long
End Type

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
   PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
   
Private strStart As String
   
Private Sub cboPrinter_Click()
    If cboPrinter.Text <> "" Then
        Set Printer = Printers(cboPrinter.ListIndex)
    End If
End Sub

Private Sub cmdCancel_Click()
    
    Dim intX As Long
    
    For intX = 0 To Printers.Count
        If Printers(intX).DeviceName = strStart Then
            Set Printer = Printers(intX)
            Exit For
        End If
    Next intX
    Unload frmPrinterSetup
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next

    Set Printer = Printers(cboPrinter.ListIndex)
    intCopies = updCopies.Value
    blnPrintAll = optAll.Value
    
    Unload frmPrinterSetup
End Sub

Private Sub cmdSetup_Click()
    Call DisplayPrinterProperties(cboPrinter.Text)
End Sub

Private Sub Form_Activate()
    Dim ptrPrinter As Printer
    Dim intX As Long
    
    cboPrinter.Clear
    
    updCopies.Value = intCopies
    optAll.Value = blnPrintAll
    optSelection.Value = Not blnPrintAll
    
    strStart = Printer.DeviceName
    
    For Each ptrPrinter In Printers
        cboPrinter.AddItem ptrPrinter.DeviceName
    Next ptrPrinter
    
    For intX = 0 To cboPrinter.ListCount - 1
        If cboPrinter.List(intX) = strStart Then
            cboPrinter.ListIndex = intX
            Exit For
        End If
    Next intX
End Sub

Public Function DisplayPrinterProperties(DeviceName As String) _
  As Boolean
         
'PURPOSE:  Displays the property sheet for the printer
'Specified by Device Name

'PARAMETER: DeviceName: DeviceName of Printer to
'Display Properties of

'EXAMPLE USAGE: DisplayPrinterProperties Printer.DeviceName

'NOTES: As Written, you must put this function into a form
'module. To put into a .bas or .cls module, add a parameter for
'the form or the form's hwnd.

On Error GoTo ErrorHandler
Dim lAns As Long, hPrinter As Long
Dim typPD As PRINTER_DEFAULTS

typPD.pDatatype = 0
typPD.pDesiredAccess = PRINTER_ALL_ACCESS
typPD.pDevMode = 0
lAns = OpenPrinter(Printer.DeviceName, hPrinter, typPD)
If lAns <> 0 Then
    lAns = PrinterProperties(Me.hwnd, hPrinter)
    DisplayPrinterProperties = lAns <> 0
End If

ErrorHandler:
If hPrinter <> 0 Then ClosePrinter hPrinter
    
End Function

