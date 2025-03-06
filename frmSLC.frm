VERSION 5.00
Begin VB.Form frmSLC 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SLC Sequence"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkcustWord 
      Caption         =   "3 Digit Word Format"
      Height          =   195
      Left            =   3240
      TabIndex        =   20
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   22
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      TabIndex        =   14
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox chkWord 
      Caption         =   "Show Word"
      Height          =   195
      Left            =   3240
      TabIndex        =   19
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkLeadZero 
      Caption         =   "Leading Zero"
      Height          =   195
      Left            =   3240
      TabIndex        =   18
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkDecimal 
      Caption         =   "Use Decimal"
      Height          =   195
      Left            =   3240
      TabIndex        =   16
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkColon 
      Caption         =   "Use Colon"
      Height          =   195
      Left            =   3240
      TabIndex        =   15
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox chkSlash 
      Caption         =   "Use Slash"
      Height          =   195
      Left            =   3240
      TabIndex        =   17
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.OptionButton optSlot 
      Caption         =   "Increment Slot"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton optWord 
      Caption         =   "Increment Word"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Frame fraIO 
      Caption         =   "Input / Output"
      Height          =   1215
      Left            =   3120
      TabIndex        =   24
      Top             =   240
      Width           =   1335
      Begin VB.OptionButton optOutputs 
         Caption         =   "Outputs"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optInputs 
         Caption         =   "Inputs"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame fraAddress 
      Caption         =   "Address"
      Height          =   1455
      Left            =   120
      TabIndex        =   23
      Top             =   240
      Width           =   2775
      Begin VB.TextBox txt3StopBit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "00"
         Top             =   840
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt3StopWord 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "000"
         Top             =   840
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txt3StartBit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "00"
         Top             =   360
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt3StartWord 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "000"
         Top             =   360
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtStopBit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "00"
         Top             =   840
         Width           =   285
      End
      Begin VB.TextBox txtStartBit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "00"
         Top             =   360
         Width           =   285
      End
      Begin VB.TextBox txtStopWord 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "00"
         Top             =   840
         Width           =   285
      End
      Begin VB.TextBox txtStartWord 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "00"
         Top             =   360
         Width           =   285
      End
      Begin VB.TextBox txtStopSlot 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "00"
         Top             =   840
         Width           =   285
      End
      Begin VB.TextBox txtStartSlot 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "00"
         Top             =   360
         Width           =   285
      End
      Begin VB.Label lblSlotDesc 
         AutoSize        =   -1  'True
         Caption         =   "Slot"
         Height          =   195
         Left            =   1200
         TabIndex        =   37
         Top             =   120
         Width           =   270
      End
      Begin VB.Label lbl3Decimal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   2160
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lbl3Decimal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   2160
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         Caption         =   "Word      Bit"
         Height          =   195
         Left            =   1680
         TabIndex        =   34
         Top             =   120
         Width           =   840
      End
      Begin VB.Label lblSlash 
         Alignment       =   2  'Center
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   33
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblSlash 
         Alignment       =   2  'Center
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   32
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblDecimal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   1515
         TabIndex        =   31
         Top             =   960
         Width           =   105
      End
      Begin VB.Label lblDecimal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1515
         TabIndex        =   30
         Top             =   480
         Width           =   105
      End
      Begin VB.Label lblIO 
         Alignment       =   1  'Right Justify
         Caption         =   "I:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   29
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblIO 
         Alignment       =   1  'Right Justify
         Caption         =   "I:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   28
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblStop 
         Alignment       =   1  'Right Justify
         Caption         =   "Stop:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblStart 
         Alignment       =   1  'Right Justify
         Caption         =   "Start:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Label lblQty 
      Caption         =   "Quantity:"
      Height          =   255
      Left            =   840
      TabIndex        =   25
      Top             =   2760
      Width           =   735
   End
End
Attribute VB_Name = "frmSLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intWord1 As Long
Dim intWord2 As Long
Dim intSlot1 As Long
Dim intSlot2 As Long
Dim intBit1 As Long
Dim intBit2 As Long
Dim strFormat As String



Private Sub chkColon_Click()
' This subroutine updates the example to hide/show the colon.

    If chkColon.Value = 1 Then
        lblIO(0).Caption = lblIO(0).Caption & ":"
        lblIO(1).Caption = lblIO(1).Caption & ":"
    Else
        lblIO(0).Caption = Left$(lblIO(0).Caption, 1)
        lblIO(1).Caption = Left$(lblIO(1).Caption, 1)
    End If
End Sub

Private Sub chkColon_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub chkcustWord_Click()
    
    If chkcustWord.Value = 1 Then
        
        ' Show all of the 3 digit word options
        lblIO(0).Left = 1320
        lblIO(1).Left = 1320
        chkColon.Value = 0
        lblSlotDesc.Visible = False
        txt3StartWord.Visible = True
        txt3StartBit.Visible = True
        txt3StopWord.Visible = True
        txt3StopBit.Visible = True
        lbl3Decimal(0).Visible = True
        lbl3Decimal(1).Visible = True
        optWord.Value = True                ' Allows us to increment the words
        optSlot.Enabled = False              ' Denies access to slot options
        chkDecimal.Enabled = False
        chkSlash.Enabled = False
        chkLeadZero.Enabled = False
                
        ' Hide the two digit word options
        txtStartSlot.Visible = False
        txtStopSlot.Visible = False
        lblDecimal(0).Visible = False
        lblDecimal(1).Visible = False
        txtStartWord.Visible = False
        txtStartBit.Visible = False
        lblSlash(0).Visible = False
        lblSlash(1).Visible = False
        txtStopWord.Visible = False
        txtStopBit.Visible = False
        
        txt3StartWord.SetFocus
        
    Else
    
         ' Hide all of the 3 digit word options
        lblIO(0).Left = 840
        lblIO(1).Left = 840
        chkColon.Value = 1
        lblSlotDesc.Visible = True
        txt3StartWord.Visible = False
        txt3StartBit.Visible = False
        txt3StopWord.Visible = False
        txt3StopBit.Visible = False
        lbl3Decimal(0).Visible = False
        lbl3Decimal(1).Visible = False
        optWord.Value = False
        optSlot.Enabled = True
        chkDecimal.Enabled = True
        chkSlash.Enabled = True
        chkLeadZero.Enabled = True

                
        ' Show the two digit word options
        txtStartSlot.Visible = True
        txtStopSlot.Visible = True
        lblDecimal(0).Visible = True
        lblDecimal(1).Visible = True
        txtStartWord.Visible = True
        txtStartBit.Visible = True
        lblSlash(0).Visible = True
        lblSlash(1).Visible = True
        txtStopWord.Visible = True
        txtStopBit.Visible = True
        txtStartWord.SetFocus

    End If
            
    
End Sub

Private Sub chkDecimal_Click()
' This subroutine hides/shows the decimals in the example.

    lblDecimal(0).Visible = Not lblDecimal(0).Visible
    lblDecimal(1).Visible = Not lblDecimal(1).Visible
End Sub

Private Sub chkDecimal_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub chkLeadZero_Click()
' This subroutine updates the example by using the leading
' zeros or not - depending on if they are selected or not.

    If chkLeadZero.Value = 1 Then
        strFormat = "00"
    Else
        strFormat = "0"
    End If
    txtStartSlot.Text = Format(Val(txtStartSlot.Text), strFormat)
    txtStopSlot.Text = Format(Val(txtStopSlot.Text), strFormat)
    txtStartWord.Text = Format(Val(txtStartWord.Text), strFormat)
    txtStopWord.Text = Format(Val(txtStopWord.Text), strFormat)
    txtStartBit.Text = Format(Val(txtStartBit.Text), strFormat)
    txtStopBit.Text = Format(Val(txtStopBit.Text), strFormat)
End Sub

Private Sub chkLeadZero_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub chkSlash_Click()
' This subroutine will update the example to show/hide slashes.

    lblSlash(0).Visible = Not lblSlash(0).Visible
    lblSlash(1).Visible = Not lblSlash(1).Visible
End Sub

Private Sub chkSlash_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub chkWord_Click()
' This subroutine will update the example to show/hide the
' word portion of the address.
    
    txtStartWord.Enabled = Not txtStartWord.Enabled
    txtStopWord.Enabled = Not txtStopWord.Enabled
    'chkDecimal.Enabled = Not chkDecimal.Enabled
    'lblDecimal(0).Enabled = Not lblDecimal(0).Enabled
    'lblDecimal(1).Enabled = Not lblDecimal(1).Enabled
End Sub

Private Sub chkWord_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub cmdCancel_Click()
' This subroutine closes out the SLC sequence dialog box.

    Unload frmSLC
End Sub

Private Sub cmdOK_Click()
' This subroutine generates the SLC sequence and adds the
' sequence to the label list.

    If chkcustWord = 1 Then
        ' use the custom format
        intSlot1 = 0        ' Just setting a default value
        intSlot2 = 1        ' Just setting a default value
        intWord1 = Val(txt3StartWord.Text)
        intWord2 = Val(txt3StopWord.Text)
        intBit1 = Val(txt3StartBit.Text)
        intBit2 = Val(txt3StopBit.Text)
    Else
        ' Use the regualar format
        intSlot1 = Val(txtStartSlot.Text)
        intSlot2 = Val(txtStopSlot.Text)
        intWord1 = Val(txtStartWord.Text)
        intWord2 = Val(txtStopWord.Text)
        intBit1 = Val(txtStartBit.Text)
        intBit2 = Val(txtStopBit.Text)
    End If

    ' Check to see that the start value is less than the stop value.
    If (intSlot1 = intSlot2 And intBit1 >= intBit2) _
    Or (intWord1 > intWord2 And optWord.Value = True) _
    Or (intSlot1 > intSlot2 And optSlot.Value = True) Then
        ' Notify user that the start value is larger than the stop value.
        MsgBox "Stopping address must be higher than starting address!", vbOKOnly + vbInformation, "SLC Sequence"
                
        If chkcustWord.Value = 1 Then
            txt3StartWord.SetFocus
        Else
            txtStartSlot.SetFocus
        End If
            
        Exit Sub
    End If
    
    frmSLC.MousePointer = 11
    
    ' Find out if we're incrementing slots or words
    If optWord.Value = True Then
        ' Use incrementing word sequence
        subWords
    Else
        ' Use incrementing slot sequence
        subSlots
    End If
    
    frmSLC.MousePointer = 0
    frmMain.lsvLabels.ListItems.Item(frmMain.lsvLabels.ListItems.Count).EnsureVisible
    frmMain.lsvLabels.Refresh
    frmMain.TotalQty
    
    ' Label list has changed and hasn't been saved yet
    blnSaved = False
    
    ' Close SLC Sequence dialog box
    Unload frmSLC
End Sub

Private Sub Form_Load()
' This subroutine sets up default values and variables.

    txtQty.Text = intDefaultQty
    strFormat = "00"
End Sub


Private Sub optInputs_Click()
' This subroutine updates the example to show all inputs.

    If chkColon.Value = 1 Then
        lblIO(0).Caption = "I:"
        lblIO(1).Caption = "I:"
    Else
        lblIO(0).Caption = "I"
        lblIO(1).Caption = "I"
    End If
End Sub

Private Sub optInputs_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub optOutputs_Click()
' This subroutine updates the example to show all outputs.

    If chkColon.Value = 1 Then
        lblIO(0).Caption = "O:"
        lblIO(1).Caption = "O:"
    Else
        lblIO(0).Caption = "O"
        lblIO(1).Caption = "O"
    End If
End Sub

Private Sub optOutputs_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub optSlot_Click()
' This subroutine updates the example to show that the word value doesn't change.
    
    chkWord.Enabled = True
    txtStopWord.Text = txtStartWord.Text
End Sub

Private Sub optSlot_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub optWord_Click()
' This subroutine updates the example to show that the slot value doesn't change.
    chkWord.Enabled = False
    chkWord.Value = 1
    txtStopSlot.Text = txtStartSlot.Text
End Sub

Private Sub optWord_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub




Private Sub txt3StartBit_GotFocus()
    SelectText
End Sub

Private Sub txt3StartBit_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the StartBit text box.  If the user presses
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

Private Sub txt3StartBit_LostFocus()
' This subroutine makes sure the value in the text box
' is valid and in the correct format.

    If Val(txt3StartBit.Text) < 0 Then
        txtStartBit.Text = "0"
    ElseIf Val(txt3StartBit.Text) > 15 Then
        txtStartBit.Text = "15"
    End If
    txt3StartBit.Text = Format(Val(txt3StartBit.Text), strFormat)
End Sub


Private Sub txt3StartWord_GotFocus()
    SelectText
End Sub

Private Sub txt3StartWord_KeyPress(KeyAscii As Integer)
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

Private Sub txt3StartWord_LostFocus()
' This subroutine checks for a valid entry and formats
' the text correctly.

    If Val(txt3StartWord.Text) < 0 Then
        txt3StartWord.Text = "000"
    'ElseIf Val(txtStartWord.Text) > 31 Then
    '    txtStartWord.Text = "31"
    End If
    
    If chkcustWord.Value = 1 Then
        strFormat = "000"
    End If
        
    txt3StartWord.Text = Format(Val(txt3StartWord.Text), strFormat)

End Sub

Private Sub txt3StopBit_GotFocus()
    SelectText
End Sub

Private Sub txt3StopBit_KeyPress(KeyAscii As Integer)
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

Private Sub txt3StopBit_LostFocus()
If Val(txt3StopBit.Text) < 0 Then
        txt3StopBit.Text = "0"
    ElseIf Val(txt3StopBit.Text) > 15 Then
        txt3StopBit.Text = "15"
    End If
    txt3StopBit.Text = Format(Val(txt3StopBit.Text), strFormat)
End Sub


Private Sub txt3StopWord_GotFocus()
    SelectText
End Sub

Private Sub txt3StopWord_KeyPress(KeyAscii As Integer)
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

Private Sub txt3StopWord_LostFocus()
' This subroutine checks for a valid entry and formats
' the text correctly.

    If Val(txt3StopWord.Text) < 0 Then
        txt3StopWord.Text = "000"
    'ElseIf Val(txtStartWord.Text) > 31 Then
    '    txtStartWord.Text = "31"
    End If
    
    If chkcustWord.Value = 1 Then
        strFormat = "000"
    End If
        
    txt3StopWord.Text = Format(Val(txt3StopWord.Text), strFormat)
End Sub

Private Sub txtQty_GotFocus()
' This subroutine selects all the text in the text box.0
    
    SelectText
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
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

Private Sub txtQty_LostFocus()
' This subroutine makes sure there is a valid quantity in the
' quantity text box.

    If Val(txtQty.Text) < 1 Then
        txtQty.Text = intDefaultQty
    End If
End Sub

Private Sub subWords()
' This subroutine generates a SLC sequence by incrementing
' the word value of each label.

    Dim intCnt1 As Long
    Dim intCnt2 As Long
    
    If chkcustWord.Value = 1 Then
    ' Formats and prints the 3 digit word labels
        For intCnt1 = intWord1 To intWord2
            For intCnt2 = intBit1 To 15
                frmMain.lsvLabels.ListItems.Add , , strWordAddress(intCnt1, intCnt2), 23, 23
                frmMain.lsvLabels.ListItems.Item(frmMain.lsvLabels.ListItems.Count).SubItems(1) = Val(txtQty.Text)
            Next intCnt2
            intBit1 = 0
        Next intCnt1
        
    Else
    ' Prints any other standard labels
        For intCnt1 = intWord1 To intWord2
            For intCnt2 = intBit1 To 15
                frmMain.lsvLabels.ListItems.Add , , strAddress(Val(txtStartSlot.Text), intCnt1, intCnt2), 23, 23
                frmMain.lsvLabels.ListItems.Item(frmMain.lsvLabels.ListItems.Count).SubItems(1) = Val(txtQty.Text)
                If intCnt1 = intWord2 And intCnt2 = intBit2 Then Exit For
            Next intCnt2
            intBit1 = 0
        Next intCnt1
        
    End If
End Sub

Private Sub subSlots()
' This subroutine generates a SLC sequence by incrementing
' the slot value of each label.

    Dim intCnt1 As Long
    Dim intCnt2 As Long
    
    
    For intCnt1 = intSlot1 To intSlot2
        For intCnt2 = intBit1 To 15
            frmMain.lsvLabels.ListItems.Add , , strAddress(intCnt1, Val(txtStartWord.Text), intCnt2), 23, 23
            frmMain.lsvLabels.ListItems.Item(frmMain.lsvLabels.ListItems.Count).SubItems(1) = Val(txtQty.Text)
            If intCnt1 = intSlot2 And intCnt2 = intBit2 Then Exit For
        Next intCnt2
        intBit1 = 0
    Next intCnt1
End Sub

Private Sub txtStartBit_GotFocus()
' This subroutine selects all the text in the text box.

    SelectText
End Sub

Private Sub txtStartBit_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the StartBit text box.  If the user presses
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

Private Sub txtStartBit_LostFocus()
' This subroutine makes sure the value in the text box
' is valid and in the correct format.

    If Val(txtStartBit.Text) < 0 Then
        txtStartBit.Text = "0"
    ElseIf Val(txtStartBit.Text) > 15 Then
        txtStartBit.Text = "15"
    End If
    txtStartBit.Text = Format(Val(txtStartBit.Text), strFormat)
End Sub

Private Sub txtStartSlot_Change()
' This subroutine keeps the stopping slot value the same
' if incrementing words is selected.

    If optWord.Value = True Then
        txtStopSlot.Text = txtStartSlot.Text
    End If
End Sub

Private Sub txtStartSlot_GotFocus()
' This subroutine selects all the text in the text box.

    SelectText
End Sub

Private Sub txtStartSlot_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the StartSlot text box.  If the user presses
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


Private Sub txtStartSlot_LostFocus()
' This subroutine checks for a valid entry and formats the text correctly.

    If Val(txtStartSlot.Text) < 0 Then
        txtStartSlot.Text = "00"
    ' ElseIf Val(txtStartSlot.Text) > 60 Then
    '    txtStartSlot.Text = "60"
    End If
    txtStartSlot.Text = Format(Val(txtStartSlot.Text), strFormat)
End Sub

Private Sub txtStartWord_Change()
' This subroutine makes sure that the word values are the same
' if increment slot is chosen.
    
    If optSlot.Value = True Then
        txtStopWord.Text = txtStartWord.Text
    End If
End Sub

Private Sub txtStartWord_GotFocus()
' This subroutine selects all the text in the text box.

    SelectText
End Sub

Private Sub txtStartWord_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the StartWord text box.  If the user presses
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

Private Sub txtStartWord_LostFocus()
' This subroutine checks for a valid entry and formats
' the text correctly.

    If Val(txtStartWord.Text) < 0 Then
        txtStartWord.Text = "0"
    ElseIf Val(txtStartWord.Text) > 31 Then
        txtStartWord.Text = "31"
    End If
    txtStartWord.Text = Format(Val(txtStartWord.Text), strFormat)
End Sub

Private Sub txtStopBit_GotFocus()
' This subroutine selects all the text in the text box.

    SelectText
End Sub

Private Sub txtStopBit_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the StopBit text box.  If the user presses
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

Private Sub txtStopBit_LostFocus()
' This subroutine checks for a valid entry and correctly
' formats the value.
    
    If Val(txtStopBit.Text) < 0 Then
        txtStopBit.Text = "0"
    ElseIf Val(txtStopBit.Text) > 15 Then
        txtStopBit.Text = "15"
    End If
    txtStopBit.Text = Format(Val(txtStopBit.Text), strFormat)
End Sub

Private Sub txtStopSlot_Change()
' This subroutine keeps the slot values the same if
' incrementing words is chosen.

    If optWord.Value = True Then
        txtStartSlot.Text = txtStopSlot.Text
    End If
End Sub

Private Sub txtStopSlot_GotFocus()
' This subroutine selects all the text in the text box.

    SelectText
End Sub

Private Sub txtStopSlot_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the StopSlot text box.  If the user presses
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

Private Sub txtStopSlot_LostFocus()
' This subroutine checks for a valid entry and makes sure
' the text is formatted correctly.
    
    If Val(txtStopSlot.Text) < 0 Then
        txtStopSlot.Text = "0"
 '   ElseIf Val(txtStopSlot.Text) > 60 Then
   '     txtStopSlot.Text = "60"
    End If
    txtStopSlot.Text = Format(Val(txtStopSlot.Text), strFormat)
End Sub

Private Sub txtStopWord_Change()
' This subroutine makes sure the word values are the same
' if incrementing slots is chosen.

    If optSlot.Value = True Then
        txtStartWord.Text = txtStopWord.Text
    End If
End Sub

Private Sub txtStopWord_GotFocus()
' This subroutine selects all the text in the text box.

    SelectText
End Sub

Private Sub txtStopWord_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the StopWord text box.  If the user presses
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

Private Sub txtStopWord_LostFocus()
' This subroutine makes sure the entry is valid and it
' puts the text into the correct format.
    
    If Val(txtStopWord.Text) < 0 Then
        txtStopWord.Text = "0"
    ElseIf Val(txtStopWord.Text) > 31 Then
        txtStopWord.Text = "31"
    End If
    txtStopWord.Text = Format(Val(txtStopWord.Text), strFormat)
End Sub

Private Function strWordAddress(ByVal intWord As Long, ByVal intbit As Long) As String
' This function formats the three digit word address
    
    Dim strWordFormat As String
    
    ' Check for input or output
    If optInputs.Value = True Then
        strWordAddress = "I"
    Else
        strWordAddress = "O"
    End If
    
    ' Check to see if colon is visible or not
    If chkColon.Value = 1 Then
        ' Add colon to string
        strWordAddress = strWordAddress & ":"
    End If

    ' Add word portion to the string
    strWordFormat = "000"
    strWordAddress = strWordAddress & Format(intWord, strWordFormat)
    
    ' Add the decimal point to the string
    strWordAddress = strWordAddress & "."
     
    ' Add bit value to end of string
    strWordFormat = "00"
    strWordAddress = strWordAddress & Format(intbit, strWordFormat)
    
End Function

Private Function strAddress(ByVal intSlot As Long, ByVal intWord As Long, ByVal intbit As Long) As String
' This function returns a correctly formatted address
' according to all the paramaters and options selected.

    ' Check if input or output
    If optInputs.Value = True Then
        ' Start string with Input "I"
        strAddress = "I"
    Else
        ' Start string with Output "O"
        strAddress = "O"
    End If
    
    ' Check to see if colon is visible or not
    If chkColon.Value = 1 Then
        ' Add colon to string
        strAddress = strAddress & ":"
    End If
    
    ' Add slot portion to string
    strFormat = "00"
    strAddress = strAddress & Format(intSlot, strFormat)
    
        
    ' Check to see if the word portion is visible
    If chkWord.Value = 1 Then
        ' Check to see if decimal point is visible
        If chkDecimal.Value = 1 Then
            ' Add decimal to string
            strAddress = strAddress & "."
        End If
        ' Add word portion to string
        strFormat = "00"
        strAddress = strAddress & Format(intWord, strFormat)
    Else
        If chkDecimal.Value = 1 Then
            ' Add decimal to string
            strAddress = strAddress & "."
        End If
        
    End If
    
'    ' Check to see if the word portion is visible
'    If chkWord.Value = 1 Then
'        ' Check to see if decimal point is visible
'        If chkDecimal.Value = 1 Then
'            ' Add decimal to string
'            strAddress = strAddress & "."
'        End If
'
'        ' Add word portion to string
'        strAddress = strAddress & Format(intWord, strFormat)
'    End If
    
    ' Check to see if the slash is visible or not
    If chkSlash.Value = 1 Then
        ' Add slash to string
        strAddress = strAddress & "/"
    End If
    
    ' Add bit value to end of string
    strFormat = "00"
    strAddress = strAddress & Format(intbit, strFormat)
End Function
