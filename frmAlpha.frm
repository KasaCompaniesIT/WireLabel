VERSION 5.00
Begin VB.Form frmAlpha 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alphabetic Sequence"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtQtyEach 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtSuffix 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtPrefix 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtStop 
      Height          =   285
      Left            =   600
      MaxLength       =   1
      TabIndex        =   6
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Left            =   600
      MaxLength       =   1
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblQtyEach 
      Caption         =   "Quantity of Each:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblSuffix 
      Caption         =   "Suffix:"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblPrefix 
      Caption         =   "Prefix:"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblStop 
      Caption         =   "Stop:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblStart 
      Caption         =   "Start:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmAlpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
' This subroutine closes the alphabetic sequence dialog box.
    Unload frmAlpha
End Sub

Private Sub cmdOK_Click()
' This subroutine makes the alphabetic sequence and adds each
' item in it to the label list, and then closes the alphabetic
' dialog box.

    Dim intTemp As Long
    Dim intStart As Long
    Dim intStop As Long
    Dim strTemp As String
        
    ' Set the starting and stopping letters by reading the ASCII
    ' values of each character
    intStart = Asc(Left$(txtStart.Text, 1))
    intStop = Asc(Left$(txtStop.Text, 1))
    
    ' Check to see that the starting letter is before the ending
    ' letter in the alphabet.
    If intStart >= intStop Then
        MsgBox "Starting character must be lower than Ending character!", vbOKOnly + vbInformation, "Alphabetic Sequence"
        txtStart.SetFocus
        Exit Sub
    End If
    
    ' Check to make sure that the user entered a valid quantity
    If Val(txtQtyEach.Text) < 1 Then
        MsgBox "Must enter a valid quantity!", vbOKOnly + vbInformation, "Decimal Sequence"
        txtQtyEach.SetFocus
        Exit Sub
    End If
    
    ' Loop through alphabet from start to stop adding each
    ' item in the loop to the label list.
    For intTemp = intStart To intStop
        strTemp = Trim$(txtPrefix.Text) & Chr(intTemp) & Trim$(txtSuffix.Text)
        frmMain.lsvLabels.ListItems.Add , , strTemp, 23, 23
        frmMain.lsvLabels.ListItems.Item(frmMain.lsvLabels.ListItems.Count).SubItems(1) = Val(txtQtyEach.Text)
    Next intTemp

    ' Label list has changed, and is not saved yet
    blnSaved = False
    
    ' Refresh label list and re-total the totals in the status bar
    frmMain.lsvLabels.Refresh
    frmMain.lsvLabels.ListItems.Item(frmMain.lsvLabels.ListItems.Count).EnsureVisible
    frmMain.TotalQty
    
    ' Close the alphabetic sequence dialog box
    Unload frmAlpha
End Sub

Private Sub Form_Activate()
' This subroutine sets the quantity to the default quantity
    txtQtyEach.Text = intDefaultQty
End Sub

Private Sub Form_Load()
' This subroutine sets the quantity to the default quantity
    
    txtQtyEach.Text = intDefaultQty
End Sub

Private Sub txtPrefix_GotFocus()
' This subroutine selects all the text in the text box.
    
    SelectText
End Sub

Private Sub txtPrefix_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtQtyEach_GotFocus()
' This subroutine selects all the text in the text box.
    SelectText
End Sub

Private Sub txtQtyEach_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the quantity text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyLeft _
    And KeyAscii <> vbKeyRight And KeyAscii <> vbKeyReturn _
    And KeyAscii <> vbKeyTab Then
        KeyAscii = 0
    End If
        
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtStart_GotFocus()
' This subroutine selects all the text in the text box.
    
    SelectText
End Sub

Private Sub txtStart_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a numeric
' character in the Start character text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKeyA Or KeyAscii > vbKeyZ) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyLeft _
    And KeyAscii <> vbKeyRight And KeyAscii <> vbKeyReturn _
    And KeyAscii <> vbKeyTab Then
        KeyAscii = 0
    End If
        
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtStop_GotFocus()
' This subroutine selects all the text in the text box.
    
    SelectText
End Sub

Private Sub txtStop_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a numeric
' character in the Start character text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKeyA Or KeyAscii > vbKeyZ) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyLeft _
    And KeyAscii <> vbKeyRight And KeyAscii <> vbKeyReturn _
    And KeyAscii <> vbKeyTab Then
        KeyAscii = 0
    End If
        
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtSuffix_GotFocus()
' This subroutine selects all the text in the text box.
    
    SelectText
End Sub

Private Sub txtSuffix_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub
