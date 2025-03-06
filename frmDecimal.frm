VERSION 5.00
Begin VB.Form frmDecimal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Decimal Sequence"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtIncrement 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "1"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtQtyEach 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.CheckBox chkLeadZeros 
      Caption         =   "Leading Zeros"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtSuffix 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtPrefix 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtStop 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "1"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtStart 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblIncrement 
      Caption         =   "Increment By:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblQtyEach 
      Caption         =   "Quantity of Each:"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblSuffix 
      Caption         =   "Suffix:"
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblPrefix 
      Caption         =   "Prefix:"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblStop 
      Caption         =   "Stop:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblStart 
      Caption         =   "Start:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmDecimal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkLeadZeros_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub cmdCancel_Click()
' This subroutine closes the Decimal Sequence Dialog Box

    Unload frmDecimal
End Sub

Private Sub cmdOK_Click()
' This subroutine generates the Decimal sequence that the user
' specified and adds each item to the label list.

    Dim intStart As Long
    Dim intStop As Long
    Dim intTemp As Long
    Dim strTemp As String
    Dim strAdd As String
    Dim intInc As Long
    
    ' Convert the start, stop and increment strings to integers
    intStart = Val(txtStart.Text)
    intStop = Val(txtStop.Text)
    intInc = Val(txtIncrement.Text)

    ' Check to make sure the start value is less than the stop value
    If intStart >= intStop Then
        MsgBox "Stop value must be greater than Start value.", vbOKOnly + vbInformation, "Decimal Sequence"
        txtStart.SetFocus
        Exit Sub
    End If
    
    ' Check to make sure that the difference between the start and stop values
    ' are greater than the increment value
    If (intStop - intStart) < intInc Then
        MsgBox "Increment value must be less than difference between start and stop values.", vbOKOnly + vbInformation, "Decimal Sequence"
        txtIncrement.SetFocus
        Exit Sub
    End If
    
    ' Check to make sure the user entered a quantity of at least one
    If Val(txtQtyEach.Text) < 1 Then
        MsgBox "Must enter a valid quantity!", vbOKOnly + vbInformation, "Decimal Sequence"
        txtQtyEach.SetFocus
        Exit Sub
    End If
    
    ' Setup the format for the labels
    If chkLeadZeros.Value = 1 Then
        strTemp = String$(Len(Trim$(txtStop.Text)), "0")
    Else
        strTemp = "0"
    End If
    
    ' Loop through the sequence, adding each label to the list
    For intTemp = intStart To intStop Step intInc
        strAdd = Trim$(txtPrefix.Text) & Format(intTemp, strTemp) & Trim$(txtSuffix.Text)
        frmMain.lsvLabels.ListItems.Add , , strAdd, 23, 23
        frmMain.lsvLabels.ListItems.Item(frmMain.lsvLabels.ListItems.Count).SubItems(1) = Trim$(Str$(Val(txtQtyEach.Text)))
    Next
    
    ' If we didn't end exactly on the stop value, then add the stop value to the list also.
    If ((intStop - intStart) Mod intInc) <> 0 Then
        strAdd = Trim$(txtPrefix.Text) & Format(intStop, strTemp) & Trim$(txtSuffix.Text)
        frmMain.lsvLabels.ListItems.Add , , strAdd, 23, 23
        frmMain.lsvLabels.ListItems.Item(frmMain.lsvLabels.ListItems.Count).SubItems(1) = Trim$(Str$(Val(txtQtyEach.Text)))
    End If
    
    ' Label list has changed, and it isn't saved yet.
    blnSaved = False
    
    ' Refresh the label list and update the status bar with the new totals.
    frmMain.lsvLabels.Refresh
    frmMain.lsvLabels.ListItems.Item(frmMain.lsvLabels.ListItems.Count).EnsureVisible
    frmMain.TotalQty
    
    ' Close the decimal sequence dialog box.
    Unload frmDecimal
End Sub

Private Sub Form_Activate()
' This subroutine loads the default quantity into the quantity text box.
    
    txtQtyEach.Text = intDefaultQty
End Sub

Private Sub Form_Load()
' This subroutine loads the default quantity into the quantity text box.
    
    txtQtyEach.Text = Str$(intDefaultQty)
End Sub

Private Sub txtIncrement_GotFocus()
' This subroutine selects all the text in the text box.
    
    SelectText
End Sub

Private Sub txtIncrement_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the increment text box.  If the user presses
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

Private Sub txtQtyEach_LostFocus()
' This subroutine makes sure that you put a valid integer in the quantity text box.
    
    txtQtyEach.Text = Trim$(Str$(Int(Val(txtQtyEach.Text))))
End Sub

Private Sub txtStart_GotFocus()
' This subroutine selects all the text in the text box.
    
    SelectText
End Sub

Private Sub txtStart_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the Start value text box.  If the user presses
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

Private Sub txtStart_LostFocus()
' This subroutine makes sure you put a valid integer in the start text box.

    txtStart.Text = Trim$(Str$(Int(Val(txtStart.Text))))
End Sub

Private Sub txtStop_GotFocus()
' This subroutine selects all the text in the text box.
    
    SelectText
End Sub

Private Sub txtStop_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the stop value text box.  If the user presses
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

Private Sub txtStop_LostFocus()
' This subroutine makes sure the user put a valid integer in the stop text box.
    
    txtStop.Text = Trim$(Str$(Int(Val(txtStop.Text))))
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
