VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowser 
   ClientHeight    =   5595
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   7095
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   794
      ButtonWidth     =   820
      ButtonHeight    =   794
      Style           =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   5400
      ExtentX         =   9525
      ExtentY         =   7858
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6180
      Top             =   1500
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   6000
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":042C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":070E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":09F0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbDontNavigateNow As Boolean
Private Sub Form_Load()
' This subroutine loads the main help file on the form load

    On Error Resume Next
    Me.Show
    tbToolBar.Refresh
    Form_Resize

    brwWebBrowser.Navigate App.Path & "\Help\default.htm"
End Sub

Private Sub brwWebBrowser_DownloadComplete()
' This subroutine makes sure the title bar reads "Kasa Wire Labels Help"
    
    Me.Caption = GetFormCaption(Me.Name)
End Sub

Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
' This subroutine makes sure the title bar reads "Kasa Wire Labels Help"
    
    Me.Caption = GetFormCaption(Me.Name)
End Sub

Private Sub Form_Resize()
' This subroutine resizes the actual browser window to fill
' the entire form when the window is resized, but not minimized.

    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    
    brwWebBrowser.Width = Me.ScaleWidth - 5
    brwWebBrowser.Height = Me.ScaleHeight - 500
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
' This subroutine runs the appropriate subroutine/procedure
' when the user clicks a button on the toolbar

    On Error Resume Next
     
    ' Enable the timer that checks to see if the browser is still
    ' loading the next help page.
    timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Back"                         ' User clicked Back Button
            ' Go Back
            brwWebBrowser.GoBack
        
        Case "Forward"                      ' User clicked Forward Button
            ' Go Forward
            brwWebBrowser.GoForward
        
        Case "Refresh"                      ' User clicked Refresh Button
            ' Refresh current pages
            brwWebBrowser.Refresh
        
        Case "Home"                         ' User clicked Home Button
            ' Go back to main help page
            brwWebBrowser.Navigate App.Path & "\Help\default.htm"
    End Select

End Sub


Private Sub timTimer_Timer()
' This subroutine checks to see when the pages are loading.
' If the pages are loading, display "Working..." otherwise
' display "Kasa Wire Labels Help"

    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = GetFormCaption(Me.Name)
    Else
        Me.Caption = "Working..."
    End If
End Sub


