VERSION 5.00
Begin VB.Form frmPrintErrorInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPrintError 
      Height          =   2655
      Index           =   484
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmPrintErrorInfo.frx":0000
      Top             =   0
      Width           =   8595
   End
   Begin VB.TextBox txtPrintError 
      Height          =   2655
      Index           =   482
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmPrintErrorInfo.frx":00B6
      Top             =   60
      Width           =   8595
   End
   Begin VB.TextBox txtPrintError 
      Height          =   2655
      Index           =   483
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmPrintErrorInfo.frx":032D
      Top             =   0
      Width           =   8595
   End
   Begin VB.TextBox txtPrintError 
      Height          =   2655
      Index           =   396
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmPrintErrorInfo.frx":0414
      Top             =   60
      Width           =   8595
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   7470
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrintErrorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    HidePrinterErrors
End Sub

Private Sub HidePrinterErrors()
    Dim c As Control
    For Each c In Me.Controls
        If TypeOf c Is TextBox Then
            If LCase(c.Name) = "txtprinterror" Then
                c.Move 60, 60
                c.Visible = False
            End If
        End If
    Next
End Sub

Public Sub DisplayPrinterErrorInfo(ByVal Number As Integer)
    HidePrinterErrors
    Select Case Number
        Case 396, 482, 483, 484
            txtPrintError(Number).Visible = True
            txtPrintError(Number).ZOrder
        Case Else
'            Debug.Assert False
            Unload Me
    End Select
End Sub
