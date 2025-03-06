VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print Status"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.Animation aniPrint 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393216
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   33
      FullHeight      =   33
   End
   Begin VB.Label lblPrint 
      Alignment       =   2  'Center
      Caption         =   "Printing..."
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
' This subroutine loads the printer animation onto the form.

    aniPrint.Open (App.Path & "\print.avi")
End Sub
