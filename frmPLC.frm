VERSION 5.00
Begin VB.Form frmPLC 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PLC Sequence"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CheckBox chkLeadZero 
      Caption         =   "Leading Zero"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkPunctuation 
      Caption         =   "Use Punctuation"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   1320
      Width           =   495
   End
   Begin VB.Frame fraAlternate 
      Caption         =   "Alternating Mode"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   2760
      TabIndex        =   16
      Top             =   1320
      Width           =   1575
      Begin VB.OptionButton optTwoSlot 
         Caption         =   "Two Slot"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optOneSlot 
         Caption         =   "One Slot"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblOutput 
         Alignment       =   1  'Right Justify
         Caption         =   "O:000/00"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblInput 
         Alignment       =   1  'Right Justify
         Caption         =   "I:000/00"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame fraIO 
      Caption         =   "Input / Output"
      Height          =   1095
      Left            =   2760
      TabIndex        =   15
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton optAlternate 
         Caption         =   "Alternating"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optOutputs 
         Caption         =   "Outputs"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optInputs 
         Caption         =   "Inputs"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame fraAddress 
      Caption         =   "Address"
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtStopBit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "00"
         Top             =   720
         Width           =   285
      End
      Begin VB.TextBox txtStartBit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "00"
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txtStopRack 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "000"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtStartRack 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "000"
         Top             =   240
         Width           =   375
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
         Left            =   1800
         TabIndex        =   23
         Top             =   720
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
         Left            =   1800
         TabIndex        =   22
         Top             =   240
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
         Index           =   1
         Left            =   1080
         TabIndex        =   21
         Top             =   720
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
         Left            =   1080
         TabIndex        =   20
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblStop 
         Alignment       =   1  'Right Justify
         Caption         =   "Stop :"
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
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblStart 
         Alignment       =   1  'Right Justify
         Caption         =   "Start :"
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
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label lblQty 
      Caption         =   "Quantity:"
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   1320
      Width           =   615
   End
End
Attribute VB_Name = "frmPLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intRack1 As Long
Dim intRack2 As Long
Dim intBit1 As Long
Dim intBit2 As Long
Dim strFrmt As String
Dim blnPunct As Boolean

Private Sub chkLeadZero_Click()
' This subroutine updates the example to reflect the leading
' zeros if they are selected.

    If chkLeadZero.Value = 0 Then
        strFrmt = "00"
    Else
        strFrmt = "000"
    End If
    txtStartRack.Text = Format(Val(txtStartRack.Text), strFrmt)
    txtStopRack.Text = Format(Val(txtStopRack.Text), strFrmt)
End Sub

Private Sub chkLeadZero_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub chkPunctuation_Click()
' This subroutine updates the example to show punctuation
' if it is selected.

    If chkPunctuation.Value = 0 Then
        ' Hide punctuation
        lblIO(0).Caption = Left$(lblIO(0).Caption, 1)
        lblIO(1).Caption = Left$(lblIO(1).Caption, 1)
        lblSlash(0).Caption = ""
        lblSlash(1).Caption = ""
        blnPunct = False
    Else
        ' Show punctuation
        lblIO(0).Caption = lblIO(0).Caption & ":"
        lblIO(1).Caption = lblIO(1).Caption & ":"
        lblSlash(0).Caption = "/"
        lblSlash(1).Caption = "/"
        blnPunct = True
    End If
End Sub

Private Sub chkPunctuation_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub cmdCancel_Click()
' This subroutine closes out the PLC sequence dialog box

    Unload frmPLC
End Sub

Private Sub cmdOK_Click()
' This subroutine generates a sequence of PLC addresses using
' the paramaters and options specified on the dialog box.

    ' Check to see if stopping value is greater than starting value
    If (Val(txtStartRack.Text) = Val(txtStopRack.Text) _
    And Val(txtStartBit.Text) >= (txtStopBit.Text)) _
    Or (Val(txtStartRack.Text) > Val(txtStopRack.Text)) Then
        ' Stop value is smaller than start value - notify user
        MsgBox "Starting values must be less than Stopping Values!", vbOKOnly + vbInformation, "PLC Sequence"
        txtStartRack.SetFocus
        Exit Sub
    End If
    
    frmPLC.MousePointer = 11
    
    ' Convert strings to integers
    intRack1 = Val(txtStartRack.Text)
    intRack2 = Val(txtStopRack.Text)
    intBit1 = Val(txtStartBit.Text)
    intBit2 = Val(txtStopBit.Text)

    ' Convert values to octal
    intRack1 = "&O" & intRack1
    intRack2 = "&O" & intRack2
    intBit1 = "&O" & intBit1
    intBit2 = "&O" & intBit2
        
    ' Convert octal values to decimal format
    intRack1 = Val(intRack1)
    intRack2 = Val(intRack2)
    intBit1 = Val(intBit1)
    intBit2 = Val(intBit2)
    
    ' Generate sequence
    If optInputs.Value = True Then
        ' All inputs
        subNormal True
    ElseIf optOutputs.Value = True Then
        ' All outputs
        subNormal False
    ElseIf optAlternate.Value = True Then
        ' Alternating inputs and outputs
        If optOneSlot.Value = True Then
            ' One slot addressing
            subAlternateOne
        Else
            ' Two slot addressing
            subAlternateTwo
        End If
    End If
    
    ' Refresh label list and re-total quantities
    frmMain.lsvLabels.Refresh
    frmMain.lsvLabels.ListItems.Item(frmMain.lsvLabels.ListItems.Count).EnsureVisible
    frmMain.TotalQty
    
    ' List has changed and isn't saved yet.
    blnSaved = False
    frmPLC.MousePointer = 0
    
    ' Close out the PLC sequence dialog box
    Unload frmPLC
End Sub

Private Sub Form_Load()
' This subroutine sets the default quantity for the sequence
' to the global default quantity, and sets the defaults for
' other values on the form.

    txtQty.Text = intDefaultQty
    strFrmt = "000"
    blnPunct = True
End Sub

Private Sub optAlternate_Click()
' This subroutine updates the examples for alternating addressing.
' It also enables the user to select one or two slot addressing.
    
    If chkPunctuation.Value = 1 Then
        lblIO(0).Caption = "I:"
        lblIO(1).Caption = "O:"
    Else
        lblIO(0).Caption = "I"
        lblIO(1).Caption = "O"
    End If
    
    fraAlternate.Enabled = True
    optOneSlot.Enabled = True
    optTwoSlot.Enabled = True
    lblInput.Enabled = True
    lblOutput.Enabled = True
End Sub

Private Sub optAlternate_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub optInputs_Click()
' This subroutine updates the example on the screen to show
' all inputs, and prevents the user from selecting one or
' two slot addressing.

    If chkPunctuation.Value = 1 Then
        lblIO(0).Caption = "I:"
        lblIO(1).Caption = "I:"
    Else
        lblIO(0).Caption = "I"
        lblIO(1).Caption = "I"
    End If
    
    fraAlternate.Enabled = False
    optOneSlot.Enabled = False
    optTwoSlot.Enabled = False
    lblInput.Enabled = False
    lblOutput.Enabled = False
End Sub

Private Sub optInputs_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub optOneSlot_Click()
' This subroutine updates the example to show one slot addressing.

    lblOutput.Caption = "O:001/00"
End Sub

Private Sub optOneSlot_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub optOutputs_Click()
' This subroutine updates the example on the screen to show
' all outputs, and prevents the user from selecting one or
' two slot addressing.

    If chkPunctuation.Value = 1 Then
        lblIO(0).Caption = "O:"
        lblIO(1).Caption = "O:"
    Else
        lblIO(0).Caption = "O"
        lblIO(1).Caption = "O"
    End If
    
    fraAlternate.Enabled = False
    optOneSlot.Enabled = False
    optTwoSlot.Enabled = False
    lblInput.Enabled = False
    lblOutput.Enabled = False
End Sub

Private Sub optOutputs_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub optTwoSlot_Click()
' This subroutine updates the example to show two slot addressing.

    lblOutput.Caption = "O:000/00"
End Sub

Private Sub optTwoSlot_KeyPress(KeyAscii As Integer)
' If the user presses the return key, it automatically
' tabs to the next field.

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtQty_GotFocus()
' This subroutine selects all the text in the text box.

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
' This subroutine makes sure there is a quantity greater than
' one in the quantity text box.  If it isn't then it sets the
' quantity to the global default quantity.
    
    If Val(txtQty.Text) < 1 Then
        txtQty.Text = intDefaultQty
    End If
End Sub

Private Sub txtStartBit_GotFocus()
' This subroutine selects all the text in the text box.

    SelectText
End Sub

Private Sub txtStartBit_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the StartBit text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey7) _
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
' This subroutine checks to make sure the value of the Start bit is less than
' or equal to 17 when it looses focus.  If it is larger, the value is set
' to 17, otherwise, the number if formatted with leading zeros.

    If Val(txtStartBit.Text) > 17 Then
        txtStartBit.Text = "17"
    End If
    txtStartBit.Text = Format(Val(txtStartBit.Text), "00")
    
End Sub

Private Sub txtStartRack_GotFocus()
' This subroutine selects all the text in the text box.

    SelectText
End Sub

Private Sub txtStartRack_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the StartRack text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey7) _
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

Private Sub txtStartRack_LostFocus()
' This subroutine makes sure the text is in the correct format.
    
    txtStartRack.Text = Format(Val(txtStartRack.Text), strFrmt)
End Sub

Private Sub txtStopBit_GotFocus()
' This subroutine selects all the text in the text box.

    SelectText
End Sub

Private Sub txtStopBit_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the StopBit text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey7) _
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
' This subroutine makes sure the text is a valid number, and
' that it is in the correct format.
    
    If Val(txtStopBit.Text) > 17 Then
        txtStopBit.Text = "17"
    ElseIf Val(txtStopBit.Text) < 0 Then
        txtStopBit.Text = "0"
    End If
    
    txtStopBit.Text = Format(Val(txtStopBit.Text), "00")
End Sub

Private Sub txtStopRack_GotFocus()
' This subroutine selects all the text in the text box.

    SelectText
End Sub

Private Sub txtStopRack_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the StopRack text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey7) _
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

Private Sub txtStopRack_LostFocus()
' This subroutine makes sure the text is in the correct format.

    txtStopRack.Text = Format(Val(txtStopRack.Text), strFrmt)
End Sub

Private Sub subAlternateOne()
' This subroutine generates the One slot addressing sequence.

    Dim blnInput As Boolean
    Dim intCnt1 As Long
    Dim intCnt2 As Long
    
    ' Start with inputs first
    blnInput = True
    
    For intCnt1 = intRack1 To intRack2
        For intCnt2 = intBit1 To &O17
            frmMain.lsvLabels.ListItems.Add , , strAddress(blnInput, Oct$(intCnt1), Oct$(intCnt2)), 23, 23
            frmMain.lsvLabels.ListItems.Item(frmMain.lsvLabels.ListItems.Count).SubItems(1) = Val(txtQty.Text)
            If intCnt1 = intRack2 And intCnt2 = intBit2 Then
                Exit For
            End If
        Next intCnt2
        
        ' Switch between inputs and outputs
        blnInput = Not blnInput
        intBit1 = 0
    Next intCnt1
End Sub

Private Sub subAlternateTwo()
' This subroutine generates the Two slot addressing sequence.
    
    Dim intInput As Long
    Dim intCnt1 As Long
    Dim intCnt2 As Long
    
    For intCnt1 = intRack1 To intRack2
        For intInput = -1 To 0
            For intCnt2 = intBit1 To &O17
                frmMain.lsvLabels.ListItems.Add , , strAddress(intInput, Oct$(intCnt1), Oct$(intCnt2)), 23, 23
                frmMain.lsvLabels.ListItems.Item(frmMain.lsvLabels.ListItems.Count).SubItems(1) = Val(txtQty.Text)
                If intCnt1 = intRack2 And intCnt2 = intBit2 Then
                    Exit For
                End If
            Next intCnt2
        Next intInput
        intBit1 = 0
    Next intCnt1
End Sub

Private Sub subNormal(ByVal blnInput As Boolean)
' This subroutine generates a normal sequence of either all
' outputs or all inputs based on whether blnInput is true or
' false.

    Dim intCnt1 As Long
    Dim intCnt2 As Long
    
    For intCnt1 = intRack1 To intRack2
        For intCnt2 = intBit1 To &O17
            frmMain.lsvLabels.ListItems.Add , , strAddress(blnInput, Oct$(intCnt1), Oct$(intCnt2)), 23, 23
            frmMain.lsvLabels.ListItems.Item(frmMain.lsvLabels.ListItems.Count).SubItems(1) = Val(txtQty.Text)
            If intCnt1 = intRack2 And intCnt2 = intBit2 Then
                Exit For
            End If
        Next intCnt2
        intBit1 = 0
    Next intCnt1
End Sub

Private Function strAddress(ByVal blnInput As Boolean, ByVal intRack As Long, ByVal intbit As Long) As String
' This function returns a correctly formatted address.

    If blnPunct = True Then
        If blnInput = True Then
            strAddress = "I:"
        Else
            strAddress = "O:"
        End If
        strAddress = strAddress & Format(intRack, strFrmt)
        strAddress = strAddress & "/"
        strAddress = strAddress & Format(intbit, "00")
    Else
        If blnInput = True Then
            strAddress = "I"
        Else
            strAddress = "O"
        End If
        strAddress = strAddress & Format(intRack, strFrmt)
        strAddress = strAddress & Format(intbit, "00")
    End If
End Function
